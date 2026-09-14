#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务队列管理器：按页面类型（初评/重评/回单）控制任务并发数。

配置（设置-任务执行方式，matcher_config.json -> queue）：
  {storage 配置文件} queue.{kind}_concurrency:
    0 = 不限，多个同时执行
    1 = 一个一个排队（默认；大文件任务并行会内存溢出崩溃，故默认排队）
    n = 最多 n 个同时执行

  queue.memory_threshold_percent: 系统内存占用率阈值(%)，默认 75
  queue.memory_rss_limit_mb:      进程内存上限(MB)，默认 0 = 不检查

同类型任务按提交顺序 FIFO 执行；不同类型互不影响。
内存看门狗（utils/memory_guard.py）在任务出队前检查内存，
紧张时保持排队并每 5 秒重试，防止大文件任务并行导致进程崩溃。

【占用计数】活跃任务数按"线程是否真的在跑"(isRunning) 统计，
并用包装过的 run() 在线程真正结束时兜底触发重新出队：
手动停止的任务走提前 return 分支、不发自定义 finished 信号，
仅靠信号计数会永久泄漏一个占用名额，导致后续任务全部"排队中"。
"""

import threading

from extend.matcher_config import MatcherConfig
from utils.memory_guard import get_memory_status
from utils.task_state import get_running_heavy_kinds

_KINDS = ("initial", "re_review", "receipt")

# 默认并发数：大文件（300M 级）任务并行执行会内存溢出崩溃，默认改为排队串行
DEFAULT_CONCURRENCY = 1
# 内存紧张时的重试间隔（秒）
_MEMORY_RETRY_INTERVAL = 5.0
# 内存受限连续重试轮数上限：超过后强制放行队首任务并告警，
# 防止 RSS 因内存碎片不回落导致队列永久卡死（36 轮 ≈ 3 分钟）
_MAX_MEMORY_RETRY_ROUNDS = 36


class TaskQueueManager:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                inst = super().__new__(cls)
                inst._queues = {k: [] for k in _KINDS}
                inst._pump_lock = threading.RLock()
                inst._retry_timer = None
                inst._memory_blocked_rounds = 0
                # 【崩溃修复】运行中 worker 的强引用列表：
                # worker 提交后调用方的局部引用可能立即消失（或被信号回调移除），
                # QThread 对象一旦在运行中被 Python GC 销毁，Qt 会 qFatal abort
                # （0xc0000409 闪退）。此处统一持有直到线程真正结束。
                inst._running = []
                cls._instance = inst
        return cls._instance

    @staticmethod
    def _sweep_finished(workers):
        """仅移除线程真正结束的 worker；运行中的（哪怕已发完自定义 finished 信号，
        run() 可能仍在收尾）一律保留"""
        return [w for w in workers if w.isRunning()]

    @classmethod
    def get_limit(cls, kind):
        try:
            cfg = MatcherConfig.load()
            val = cfg.get("queue", {}).get(f"{kind}_concurrency", DEFAULT_CONCURRENCY)
            return max(0, int(val))
        except Exception:
            return DEFAULT_CONCURRENCY

    def _active_count(self, kind):
        """某类型正在运行的线程数（以 isRunning 为准，杜绝信号遗漏导致的计数泄漏）。
        调用方须持有 _pump_lock。"""
        self._running = self._sweep_finished(self._running)
        return sum(
            1
            for w in self._running
            if getattr(w, "_queue_kind", None) == kind and w.isRunning()
        )

    def _wrap_run(self, worker, kind):
        """包装 worker.run：线程退出（任何路径）时兜底触发重新出队。
        调用方须持有 _pump_lock。"""
        if getattr(worker, "_run_wrapped", False):
            return
        orig_run = worker.run

        def _run_wrapper():
            try:
                orig_run()
            finally:
                # 此刻线程还未真正结束（isRunning 仍为 True），不能在当前线程
                # 里直接重新出队；派生一个等待线程，待其退出后释放占用。
                threading.Thread(
                    target=self._wait_exit_and_pump,
                    args=(worker, kind),
                    daemon=True,
                ).start()

        worker.run = _run_wrapper
        worker._run_wrapped = True

    def _wait_exit_and_pump(self, worker, kind):
        """等待线程真正结束后释放占用并重新出队"""
        try:
            worker.wait(60000)
        except Exception:
            pass
        self._on_thread_exit(kind)

    def submit(self, kind, worker):
        """
        提交任务线程。立即启动返回 True；进入队列等待返回 False。
        """
        with self._pump_lock:
            self._running = self._sweep_finished(self._running)

        if kind not in _KINDS:
            with self._pump_lock:
                self._running.append(worker)
            worker.start()
            return True

        worker._queue_kind = kind
        limit = self.get_limit(kind)
        with self._pump_lock:
            self._queues[kind].append(worker)
            self._memory_blocked_rounds = 0
            self._pump(kind)
        return getattr(worker, "_queue_started", False)

    def _on_thread_exit(self, kind):
        """包装 run() 的 finally 回调：线程真正结束（含手动停止的提前 return、
        未捕获异常），释放占用并尝试出队下一个任务。"""
        with self._pump_lock:
            self._running = self._sweep_finished(self._running)
            self._memory_blocked_rounds = 0
            # 有任务结束释放内存，所有类型都可以尝试重新出队
            for k in _KINDS:
                self._pump(k)

    def _pump(self, kind):
        # 调用方须持有 _pump_lock
        if kind not in self._queues:
            return

        limit = self.get_limit(kind)
        if limit <= 0:
            # 配置为不限：把排队的全部放行（仍受大任务互斥+内存看门狗约束）
            while self._queues[kind]:
                if not self._memory_gate(kind):
                    return
                w = self._queues[kind].pop(0)
                self._start(kind, w)
            return

        while self._active_count(kind) < limit and self._queues[kind]:
            if not self._memory_gate(kind):
                return
            w = self._queues[kind].pop(0)
            self._start(kind, w)

    def _memory_gate(self, kind):
        """出队前的内存闸门 + 大任务跨类型互斥。返回 True 表示放行；
        受阻时返回 False 并安排重试，连续超限时强制放行防卡死。"""
        # 1. 大任务互斥：大型项目运行期间，其他类型任务一律排队，
        #    防止重内存任务叠加导致进程被系统终止（内存看门狗只在启动瞬间
        #    检查，拦不住启动后内存暴涨的场景）
        heavy_kinds = get_running_heavy_kinds()
        if heavy_kinds and kind not in heavy_kinds:
            self._memory_blocked_rounds += 1
            if self._memory_blocked_rounds <= 1:
                self._log_event(
                    f"[任务队列] 大型项目({'+'.join(sorted(heavy_kinds))})运行中，"
                    f"{kind} 任务暂缓，待其结束后自动开始", "WARN")
            self._schedule_memory_retry()
            return False

        # 2. 内存闸门
        ok, desc = get_memory_status()
        if ok:
            self._memory_blocked_rounds = 0
            return True
        if self._memory_blocked_rounds >= _MAX_MEMORY_RETRY_ROUNDS:
            self._memory_blocked_rounds = 0
            self._log_event(f"[任务队列] 内存持续紧张（{desc}），已达重试上限，"
                            f"强制放行队首任务以避免队列卡死", "WARN")
            return True
        self._memory_blocked_rounds += 1
        self._log_hold(desc)
        self._schedule_memory_retry()
        return False

    def _start(self, kind, worker):
        # 调用方须持有 _pump_lock
        with self._pump_lock:
            self._running = self._sweep_finished(self._running)
            self._running.append(worker)
        self._wrap_run(worker, kind)
        worker._queue_started = True
        worker.start()
        self._log_event(f"[任务队列] {kind} 任务启动 "
                        f"(活跃 {self._active_count(kind)}, 排队 {len(self._queues[kind])})")

    def _log_hold(self, desc):
        self._log_event(f"[任务队列] 内存紧张（{desc}），任务暂缓启动，"
                        f"{_MEMORY_RETRY_INTERVAL:.0f} 秒后重试", "WARN")

    def _log_event(self, message, level="DEBUG"):
        try:
            from utils.runtime_logger import RuntimeLogger
            RuntimeLogger.log(message, level)
        except Exception:
            pass

    def _schedule_memory_retry(self):
        """内存紧张时安排定时重试（全局唯一计时器，避免堆积）"""
        with self._pump_lock:
            if self._retry_timer is not None:
                return
            timer = threading.Timer(_MEMORY_RETRY_INTERVAL, self._memory_retry)
            timer.daemon = True
            self._retry_timer = timer
        timer.start()

    def _memory_retry(self):
        """定时重试：重新尝试出队所有类型的排队任务"""
        with self._pump_lock:
            self._retry_timer = None
        with self._pump_lock:
            for k in _KINDS:
                if self._queues.get(k):
                    self._pump(k)

    def queued_count(self, kind):
        with self._pump_lock:
            return len(self._queues.get(kind, []))
