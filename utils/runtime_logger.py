#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
运行日志（v3）：项目日志与应用日志分离。

架构说明
========
- 项目日志：任务线程通过 set_project() 设置【线程局部】项目名后，其日志写入
  {storage.logs}/{项目名}/{项目名}_{启动时间}.log，每个任务一个文件，并行任务
  互不干扰（v1 的全局单文件在并行时会被覆盖，已用线程局部上下文解决）。
- 应用日志：无项目上下文的日志（界面操作、启动/关闭、队列决策、看门狗等）
  统一写入 {storage.logs}/application/app_YYYYMMDD.log（按天轮转，保留30天），
  由 ApplicationLogger（stdlib logging，线程安全）落盘。
- 两类日志互不混写。

公开 API（保持向后兼容）:
  RuntimeLogger.log(msg, level)           记录日志（按线程上下文自动分流）
  RuntimeLogger.set_project(name)         设置当前线程的项目上下文
  RuntimeLogger.get_project()             读取项目名（线程局部，回退最近一次）
  RuntimeLogger.set_log_level(level)      设置全局日志级别
  RuntimeLogger.get_current_log_path(project=None)   项目/最近日志文件路径
  RuntimeLogger.get_log_dir(project=None) 项目日志目录
  RuntimeLogger.start_session() / save_session() / save_to_file() / close_file()
  RuntimeLogger.add_debug_tag(tag) / classify_level(msg, default)

模块级便捷函数（供代码内直接调用，替代 print）:
  log_debug / log_info / log_warn / log_error(*parts, sep=" ")
"""
import os
import sys
import logging
import threading
from datetime import datetime

from utils.app_logger import ApplicationLogger


def _stringify(parts, sep=" "):
    """print 语义的参数拼接：非字符串 str() 后按 sep 连接"""
    return sep.join(p if isinstance(p, str) else str(p) for p in parts)


def log_debug(*parts, sep=" "):
    """记录 DEBUG 级别日志（替代诊断类 print）"""
    RuntimeLogger.log(_stringify(parts, sep), "DEBUG")


def log_info(*parts, sep=" "):
    """记录 INFO 级别日志（替代流程类 print）"""
    RuntimeLogger.log(_stringify(parts, sep), "INFO")


def log_warn(*parts, sep=" "):
    """记录 WARN 级别日志（替代警告类 print）"""
    RuntimeLogger.log(_stringify(parts, sep), "WARN")


def log_error(*parts, sep=" "):
    """记录 ERROR 级别日志（替代错误类 print）"""
    RuntimeLogger.log(_stringify(parts, sep), "ERROR")


def _get_logs_base_dir():
    """日志根目录：优先设置中的 storage.logs，失败回退到程序目录/logs"""
    try:
        from extend.matcher_config import MatcherConfig
        base = (MatcherConfig.load().get("storage", {}) or {}).get("logs")
        if base:
            return base
    except Exception:
        pass
    return os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs"
    )


def _safe_project_name(project):
    return "".join(
        c for c in str(project or "") if c not in '<>:"/\\|?*'
    ).strip() or "默认项目"


class RuntimeLogger:
    """运行日志门面（单例，线程安全）：项目日志按项目分文件，应用日志统一落盘"""

    _instance = None

    # 重定向前的真实控制台流（打包无窗口模式可能为 None）
    _original_stdout = sys.__stdout__
    _original_stderr = sys.__stderr__

    # 日志级别：数值越大越详细，低于当前设置的日志被丢弃
    _log_level = "INFO"  # NONE, ERROR, WARN, INFO, DEBUG
    _level_map = {"NONE": 0, "ERROR": 1, "WARN": 2, "INFO": 3, "DEBUG": 4}
    _py_levels = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARN": logging.WARNING,
        "ERROR": logging.ERROR,
    }
    _level_lock = threading.Lock()

    # 线程局部：当前线程的项目上下文（并行任务互不影响）
    _local = threading.local()
    # 最近一次设置的项目名（供 path_utils 等跨线程兼容场景回退使用）
    _last_project = None

    # 项目日志文件句柄注册表：{路径: 文件对象}，跨线程共享（追加模式+写锁）
    _handles = {}
    _handles_lock = threading.Lock()
    # 项目 → 日志文件路径 注册表（供 get_current_log_path 查询）
    _project_paths = {}
    _last_log_path = None

    # ==================================================================
    # 调试标签清单
    # 凡是以这些前缀输出的内容（含经 stdout 兜底通道进来的 print），
    # 一律按 DEBUG 级别记录，受"日志级别"设置统一过滤。
    # 如有新的调试前缀，调用 RuntimeLogger.add_debug_tag("[XXX]") 即可。
    # ==================================================================
    _AUTO_DEBUG_TAGS = [
        "[DEBUG]",
        "[LOG-LEVEL-CHANGE]",
        "[COMBO-UPDATE]",
        "[ASSET-SHEET]",
        "[EXCEL-COMBO]",
        "[SHEET-CHANGED]",
        "[SIMPLE-SHEET]",
        "[SMART",            # 同时覆盖 [SMART] 与 [SMART-merged]
        "[CALL]",
        "[TRACE]",
        "[PERF]",
        ">>> ENTERING",
        "<<< EXITING",
    ]

    @classmethod
    def add_debug_tag(cls, tag: str):
        """注册新的调试前缀：以该前缀开头的输出按 DEBUG 级别过滤"""
        tag = (tag or "").strip()
        if tag and tag not in cls._AUTO_DEBUG_TAGS:
            cls._AUTO_DEBUG_TAGS.append(tag)

    @classmethod
    def classify_level(cls, message: str, default_level: str = "INFO") -> str:
        """根据消息前缀判定其实际级别：命中调试标签 → DEBUG"""
        stripped = (message or "").lstrip()
        for tag in cls._AUTO_DEBUG_TAGS:
            if stripped.startswith(tag):
                return "DEBUG"
        return default_level

    class StreamToLogger:
        """stdout/stderr 兜底重定向：捕获漏网 print 与第三方库输出，统一走级别门控"""

        def __init__(self, original_stream, level="INFO"):
            self.original_stream = original_stream  # 仅用于 flush
            self.level = level

        def write(self, buf):
            for line in buf.splitlines(True):
                msg = line.rstrip("\n")
                if not msg.strip():
                    continue
                # 命中调试标签的 print 输出按 DEBUG 过滤；日志与控制台输出
                # 统一由 RuntimeLogger.log 处理（含级别门控），避免二次转发
                actual_level = RuntimeLogger.classify_level(msg, self.level)
                RuntimeLogger.log(msg, actual_level)

        def flush(self):
            stream = self.original_stream
            if stream:
                try:
                    stream.flush()
                except Exception:
                    pass

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(RuntimeLogger, cls).__new__(cls)
            cls._setup_redirection()
        return cls._instance

    @classmethod
    def _setup_redirection(cls):
        if not isinstance(sys.stdout, cls.StreamToLogger):
            sys.stdout = cls.StreamToLogger(cls._original_stdout, "INFO")
            sys.stderr = cls.StreamToLogger(cls._original_stderr, "ERROR")

    @classmethod
    def set_log_level(cls, level: str):
        """全局设置日志级别（设置页保存后即时生效）"""
        new_level = (level or "INFO").upper()
        if new_level not in cls._level_map:
            new_level = "INFO"
        with cls._level_lock:
            old_level = cls._log_level
            cls._log_level = new_level
        if old_level != new_level:
            cls.log(f"日志级别变更: {old_level} → {new_level}", "INFO")

    @classmethod
    def set_project(cls, project_name):
        """设置当前线程的项目上下文（线程局部；并行任务互不影响）。
        每次设置都会开启该任务本次运行的项目日志文件。"""
        if project_name:
            cls._local.project = str(project_name)
            cls._local.file_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            with cls._level_lock:
                cls._last_project = str(project_name)
        else:
            cls._local.project = None
            cls._local.file_ts = None

    @classmethod
    def get_project(cls, fallback_last=True):
        """当前线程的项目名；未设置的线程（如主线程）回退到最近一次设置的项目"""
        project = getattr(cls._local, "project", None)
        if project is None and fallback_last:
            with cls._level_lock:
                project = cls._last_project
        return project

    @classmethod
    def _get_project_log_path(cls, project):
        """获取（必要时创建）当前线程上下文的项目日志文件路径"""
        file_ts = getattr(cls._local, "file_ts", None) or \
            datetime.now().strftime("%Y%m%d_%H%M%S")
        cls._local.file_ts = file_ts
        safe = _safe_project_name(project)
        base = _get_logs_base_dir()
        log_dir = os.path.join(base, safe)
        os.makedirs(log_dir, exist_ok=True)
        return os.path.join(log_dir, f"{safe}_{file_ts}.log")

    @classmethod
    def _write_project_log(cls, project, text, level):
        """写入项目日志文件（追加模式 + 全局写锁，跨线程安全）"""
        path = cls._get_project_log_path(project)
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{timestamp}] [{level}] {text}\n"
        try:
            with cls._handles_lock:
                handle = cls._handles.get(path)
                if handle is None or handle.closed:
                    handle = open(path, "a", encoding="utf-8")
                    cls._handles[path] = handle
                handle.write(line)
                handle.flush()
            cls._project_paths[project] = path
            cls._last_log_path = path
        except Exception:
            # 落盘失败不影响控制台输出，也不中断任务
            pass

    @classmethod
    def log(cls, message, level="INFO"):
        """记录日志：级别过滤 → 按线程上下文分流（项目日志 / 应用日志）→ 控制台输出"""
        level = (level or "INFO").upper()
        if level not in cls._level_map:
            level = "INFO"

        # 1. 级别过滤：低于当前设置的日志直接丢弃
        if cls._level_map[level] > cls._level_map.get(cls._log_level, 3):
            return

        msg = str(message or "")
        if not msg.strip():
            return

        # 防重入：落盘过程自身产生的输出（如 logging 异常写 stderr）不再进入本方法
        if getattr(cls._local, "in_log", False):
            return
        cls._local.in_log = True
        try:
            project = getattr(cls._local, "project", None)

            # 2. 控制台输出（重定向前真实流；打包无窗口模式为 None 时跳过）
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            console = cls._original_stderr if level == "ERROR" else cls._original_stdout
            if console:
                text = f"[{project}] {msg}" if project else msg
                try:
                    console.write(f"[{timestamp}] [{level}] {text}\n")
                except Exception:
                    pass

            # 3. 落盘分流：任务线程 → 项目日志；其余（界面/队列/系统）→ 应用日志
            if project:
                cls._write_project_log(project, msg, level)
            else:
                try:
                    ApplicationLogger.get_logger().log(
                        cls._py_levels[level], msg
                    )
                except Exception:
                    pass
        finally:
            cls._local.in_log = False

    # ------------------------------------------------------------------
    # 兼容 API
    # ------------------------------------------------------------------
    @classmethod
    def get_log_dir(cls, project_name=None):
        """项目日志目录；未指定项目时返回日志根目录（含 application/ 子目录）"""
        project = project_name or getattr(cls._local, "project", None)
        base = _get_logs_base_dir()
        if project:
            project = str(project)
            registered = cls._project_paths.get(project)
            if registered:
                return os.path.dirname(registered)
            return os.path.join(base, _safe_project_name(project))
        return base

    @classmethod
    def get_current_log_path(cls, project_name=None):
        """项目日志文件路径；未指定项目时返回最近一次写入的项目日志"""
        project = project_name or getattr(cls._local, "project", None)
        if project:
            project = str(project)
            if project in cls._project_paths:
                return cls._project_paths[project]
            # 该项目尚未写过日志
            if project_name:
                return None
        return cls._last_log_path

    @classmethod
    def close_file(cls):
        """程序退出时调用：关闭所有项目日志文件句柄"""
        with cls._handles_lock:
            for handle in cls._handles.values():
                try:
                    if not handle.closed:
                        handle.close()
                except Exception:
                    pass
            cls._handles.clear()

    @classmethod
    def clear(cls):
        """兼容保留：无内存缓冲，日志直接落盘"""

    @classmethod
    def start_session(cls):
        cls()
        cls.log("=== CosmicReviewGUI 运行开始 ===")

    @classmethod
    def save_session(cls):
        return None

    @classmethod
    def save_to_file(cls, filename, output_dir=None):
        """兼容保留：返回指定项目（或当前线程项目）的日志文件路径"""
        return cls.get_current_log_path(filename)

    @classmethod
    def get_logs(cls):
        """兼容保留：无内存缓冲，历史日志请读取对应日志文件"""
        return ""
