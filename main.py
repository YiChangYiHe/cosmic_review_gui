# -*- coding: utf-8 -*-
import sys
import os
import traceback
from utils.app_logger import ApplicationLogger


# ==========================================================
# 0. PyInstaller 打包环境兼容性修复 (核心：解决图标/字体不显示)
# ==========================================================
if getattr(sys, 'frozen', False):
    # 如果是打包后的环境
    base_path = sys._MEIPASS

    # 1. 强制指定 Qt 插件路径 (解决图标、下拉箭头、字体渲染等 UI 元素丢失问题)
    qt_plugin_path = os.path.join(base_path, 'PySide6', 'plugins')
    if os.path.exists(qt_plugin_path):
        os.environ['QT_PLUGIN_PATH'] = qt_plugin_path
        # 告诉 QCoreApplication 去哪里找插件
        from PySide6.QtCore import QCoreApplication

        QCoreApplication.addLibraryPath(qt_plugin_path)

    # 2. 将 PySide6 目录加入系统 PATH (解决部分底层 DLL 或字体加载失败的问题)
    pyside6_path = os.path.join(base_path, 'PySide6')
    if os.path.exists(pyside6_path):
        os.environ['PATH'] = pyside6_path + os.pathsep + os.environ.get('PATH', '')
else:
    # 开发环境
    base_path = os.path.abspath(os.path.dirname(__file__))

# 将项目根目录加入 sys.path，确保内部模块导入正常
sys.path.insert(0, base_path)

# ==========================================================
# 1. 核心依赖导入
# ==========================================================
try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QIcon

    from ui.main_window import CosmicMainWindow
    from utils.path_utils import get_resource_path
    from utils.runtime_logger import RuntimeLogger, log_error, log_warn
    from extend.matcher_config import MatcherConfig
except ImportError as e:
    print(f"核心组件加载失败: {e}")
    if sys.platform == "win32":
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            0,
            f"核心组件加载失败。\n\n错误: {e}\n\n请检查 PySide6 是否安装完整，或依赖是否缺失。",
            "启动失败",
            16,
        )
    sys.exit(1)


def main():
    """主程序入口"""
    # 【新增】应用用户设置的日志级别（须在写第一条日志前生效）
    try:
        RuntimeLogger.set_log_level(MatcherConfig.load().get("log_level", "INFO"))
    except Exception:
        RuntimeLogger.set_log_level("INFO")

    # 【新增】初始化全局应用日志
    ApplicationLogger.log_startup(version="1.0.0")

    # 【新增】配置全局异常处理器
    def global_exception_handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        ApplicationLogger.log_exception(exc_value, "未捕获的全局异常")

    sys.excepthook = global_exception_handler

    # 1. 设置高 DPI 缩放策略 (必须在 QApplication 实例化前)
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    # 2. 创建应用实例
    app = QApplication(sys.argv)
    app.setApplicationName("COSMIC智能评估")

    # 【新增】接管 Qt 内部消息（含 qFatal 致命错误）。Qt 的致命错误走 C 层
    # stderr，Python 的 stdout 重定向捕获不到；不接管的话闪退时没有任何线索。
    # 典型致命错误: "QThread: Destroyed while thread is still running"
    from PySide6.QtCore import qInstallMessageHandler, QtMsgType

    def _qt_message_handler(msg_type, context, message):
        try:
            # 样式表的 "Unknown property transition/transform" 警告数量巨大
            # （QSS 里写了 Qt 不认识的 CSS 动画属性），降为 DEBUG 防止刷屏
            if "Unknown property" in message:
                label = "DEBUG"
            else:
                label = {
                    QtMsgType.QtDebugMsg: "DEBUG",
                    QtMsgType.QtInfoMsg: "INFO",
                    QtMsgType.QtWarningMsg: "WARN",
                    QtMsgType.QtCriticalMsg: "ERROR",
                    QtMsgType.QtFatalMsg: "FATAL",
                }.get(msg_type, "INFO")
            loc = ""
            if context is not None and context.file:
                loc = f" ({context.file}:{context.line})"
            ApplicationLogger.log_qt_message(label, f"[Qt] {message}{loc}")
        except Exception:
            pass

    qInstallMessageHandler(_qt_message_handler)

    # 3. 设置全局样式 (Fusion 样式在不同平台上表现更一致)
    app.setStyle("Fusion")

    # 4. 设置全局默认字体
    font = QFont("Microsoft YaHei UI", 9)
    if not font.exactMatch():
        font = QFont("Microsoft YaHei", 9)
    app.setFont(font)

    # 5. 设置应用图标
    try:
        icon_path = get_resource_path("ui/logo.png")
        if not os.path.exists(icon_path):
            icon_path = get_resource_path("ui/logo.ico")
        if os.path.exists(icon_path):
            app_icon = QIcon(icon_path)
            app.setWindowIcon(app_icon)
        else:
            log_warn(f"警告: 未找到应用图标文件: {icon_path}")
    except Exception as e:
        log_error(f"设置应用图标时发生异常: {e}")
        ApplicationLogger.log_exception(e, "设置应用图标")

    # 6. 创建并显示主窗口
    try:
        window = CosmicMainWindow()
        window.show()
        ApplicationLogger.log_thread_event("MainThread", "主窗口创建成功")
    except Exception as e:
        log_error(f"主窗口初始化失败: {e}")
        log_error(traceback.format_exc())
        ApplicationLogger.log_exception(e, "主窗口初始化")
        return 1

    # 7. 启动事件循环
    try:
        exit_code = app.exec()
        ApplicationLogger.log_shutdown()
        return exit_code
    except Exception as e:
        ApplicationLogger.log_exception(e, "应用运行异常")
        return 1


if __name__ == "__main__":
    # 1. 启动全局日志会话
    RuntimeLogger.start_session()
    try:
        # 执行主程序
        exit_code = main()
        sys.exit(exit_code)
    except Exception as e:
        # 捕获未处理的顶层异常并记录到日志
        error_msg = f"程序发生严重未捕获错误:\n{str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        try:
            RuntimeLogger.log(error_msg, level="ERROR")
        except:
            pass
        sys.exit(1)
    finally:
        # 确保程序退出时保存日志并释放资源
        try:
            RuntimeLogger.save_session()
            RuntimeLogger.close_file()
        except:
            pass