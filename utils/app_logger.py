#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全局应用日志器 - 记录整个软件运行状态（按天分割）
独立于 RuntimeLogger，用于追踪多任务并发、崩溃等全局问题
"""
import os
import sys
import json
import logging
import traceback
import threading
from datetime import datetime
from pathlib import Path


# 应用根目录锚定：配置不可用时的兜底日志位置
_APP_ROOT = Path(__file__).resolve().parent.parent

# 本次运行的应用日志文件（每启动一次应用创建一个新文件）
_SESSION_FILE = None


def _resolve_log_root() -> Path:
    """应用日志根目录：优先设置中的 storage.logs，失败回退到程序目录/logs"""
    try:
        from extend.matcher_config import MatcherConfig
        base = (MatcherConfig.load().get("storage", {}) or {}).get("logs")
        if base:
            return Path(base) / "application"
    except Exception:
        pass
    return _APP_ROOT / "logs" / "application"


def get_app_log_dir() -> str:
    """本次运行的应用日志目录：{日志根}/application/{当天日期}/（按天分文件夹）"""
    d = _resolve_log_root() / datetime.now().strftime("%Y%m%d")
    d.mkdir(parents=True, exist_ok=True)
    return str(d)


def get_app_log_file_path() -> str:
    """本次运行的应用日志文件路径：每次启动创建一个新文件（含启动时间），
    进程被关闭或正常退出后该文件即完整保留本次运行的全部应用日志"""
    global _SESSION_FILE
    if _SESSION_FILE:
        return _SESSION_FILE
    day_dir = Path(get_app_log_dir())
    _SESSION_FILE = str(
        day_dir / f"app_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    return _SESSION_FILE


class ApplicationLogger:
    """全局应用日志器（单例模式，仅落盘；控制台输出由 RuntimeLogger 统一处理）"""

    _instance = None
    _lock = threading.Lock()
    _logger = None

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        # 双重检查加锁：多线程首次并发调用 get_logger 时防止重复初始化
        # （重复初始化会 clear 掉其他线程刚装好的 handler，造成窗口期日志丢失）
        # 注意：_initialized 必须在 _setup_logger 完成后才置位，否则其他线程
        # 会在初始化窗口内拿到 _logger=None，其日志被静默丢弃
        if getattr(self, '_initialized', False):
            return
        with ApplicationLogger._lock:
            if getattr(self, '_initialized', False):
                return
            self._setup_logger()
            self._initialized = True

    def _setup_logger(self):
        """配置日志系统"""
        # 创建日志目录
        get_app_log_dir()

        # 日志文件路径（按天分割）
        log_file = get_app_log_file_path()

        # 创建日志器
        self._logger = logging.getLogger("ApplicationLogger")
        self._logger.setLevel(logging.DEBUG)
        self._logger.propagate = False

        # 清除已有的 handler
        if self._logger.handlers:
            self._logger.handlers.clear()

        # 文件处理器 - 详细日志（每次运行一个独立文件，无需轮转）
        file_handler = logging.FileHandler(
            log_file,
            mode='a',
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(threadName)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))

        self._logger.addHandler(file_handler)
        # 会话头（直接写，不经 get_logger —— 初始化尚未完成，避免锁重入死锁）
        try:
            self._logger.info(
                f"========== 应用日志会话开始 | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =========="
            )
        except Exception:
            pass

    @classmethod
    def log_session_header(cls):
        """会话头：标记本次应用日志文件的开始时间"""
        try:
            cls.get_logger().info(
                f"========== 应用日志会话开始 | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =========="
            )
        except Exception:
            pass

    @classmethod
    def get_logger(cls):
        """获取日志器实例"""
        instance = cls()
        return instance._logger

    @classmethod
    def log_qt_message(cls, label, message):
        """记录 Qt 内部消息（含 qFatal 致命错误）——无论日志级别设置如何，
        Qt 致命错误是闪退的唯一现场证据，必须落盘"""
        levels = {
            "DEBUG": logging.DEBUG, "INFO": logging.INFO,
            "WARN": logging.WARNING, "ERROR": logging.ERROR,
            "FATAL": logging.ERROR,
        }
        try:
            cls.get_logger().log(levels.get(label, logging.INFO), message)
        except Exception:
            pass

    @classmethod
    def log_startup(cls, version="1.0.0"):
        """记录软件启动"""
        logger = cls.get_logger()
        logger.info("=" * 80)
        logger.info(f"🚀 应用启动 | 版本: {version}")
        logger.info(f"📅 启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"💻 系统: {os.name} | Python: {sys.version.split()[0]}")
        logger.info(f"🧵 主线程: {threading.current_thread().name}")
        logger.info("=" * 80)

    @classmethod
    def log_shutdown(cls):
        """记录软件关闭"""
        logger = cls.get_logger()
        logger.info("=" * 80)
        logger.info(f"🛑 应用关闭 | 时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f" 活跃线程数: {threading.active_count()}")
        logger.info("=" * 80)

    @classmethod
    def log_exception(cls, exc: Exception, context: str = ""):
        """记录异常"""
        logger = cls.get_logger()
        logger.error(f"❌ 异常发生 | 上下文: {context}")
        logger.error(f"异常类型: {type(exc).__name__}")
        logger.error(f"异常信息: {str(exc)}")
        logger.error("堆栈跟踪:\n" + traceback.format_exc())

    @classmethod
    def log_thread_event(cls, thread_name: str, event: str, details: str = ""):
        """记录线程事件"""
        logger = cls.get_logger()
        logger.debug(f"🧵 [{thread_name}] {event} | {details}")

    @classmethod
    def log_memory_usage(cls):
        """记录内存使用情况"""
        try:
            import psutil
            process = psutil.Process()
            memory_info = process.memory_info()
            logger = cls.get_logger()
            logger.debug(f"💾 内存使用: RSS={memory_info.rss / 1024 / 1024:.2f}MB | "
                         f"VMS={memory_info.vms / 1024 / 1024:.2f}MB | "
                         f"线程数: {threading.active_count()}")
        except ImportError:
            pass

    @classmethod
    def log_task_start(cls, task_type: str, task_id: str, details: dict = None):
        """记录任务开始"""
        logger = cls.get_logger()
        logger.info(f"▶️ 任务开始 | 类型: {task_type} | ID: {task_id}")
        if details:
            logger.debug(f"任务详情: {details}")

    @classmethod
    def log_task_end(cls, task_type: str, task_id: str, success: bool, error: str = None):
        """记录任务结束"""
        logger = cls.get_logger()
        if success:
            logger.info(f"✅ 任务完成 | 类型: {task_type} | ID: {task_id}")
        else:
            logger.error(f"❌ 任务失败 | 类型: {task_type} | ID: {task_id} | 错误: {error}")

    @classmethod
    def log_concurrent_status(cls, active_tasks: int, queued_tasks: int):
        """记录并发状态"""
        logger = cls.get_logger()
        logger.debug(f"🔄 并发状态 | 活跃任务: {active_tasks} | 排队任务: {queued_tasks} | "
                     f"总线程: {threading.active_count()}")


# 便捷的全局访问函数
def get_app_logger():
    """获取应用日志器"""
    return ApplicationLogger.get_logger()


def log_exception(exc: Exception, context: str = ""):
    """记录异常"""
    ApplicationLogger.log_exception(exc, context)


def log_task_start(task_type: str, task_id: str, details: dict = None):
    """记录任务开始"""
    ApplicationLogger.log_task_start(task_type, task_id, details)


def log_task_end(task_type: str, task_id: str, success: bool, error: str = None):
    """记录任务结束"""
    ApplicationLogger.log_task_end(task_type, task_id, success, error)