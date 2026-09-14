# utils/notification_manager.py
import os
from PySide6.QtWidgets import QApplication, QLabel, QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QSystemTrayIcon
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon
from extend.matcher_config import MatcherConfig
from utils.path_utils import get_resource_path
from utils.runtime_logger import log_error


class NotificationManager:
    """全局通知管理器"""

    _tray_instance = None  # 保存托盘图标实例，防止被垃圾回收
    _alert_timer = None  # 闪烁定时器

    @classmethod
    def notify(cls, title, message, parent_window, has_issue=False, on_view_callback=None):
        """
        统一通知入口
        :param title: 通知标题
        :param message: 通知内容
        :param parent_window: 主窗口实例 (用于定位 Toast 和闪烁任务栏)
        :param has_issue: 是否有异常 (影响图标和颜色)
        :param on_view_callback: 点击“查看结果”时的回调函数
        """
        try:
            config = MatcherConfig.load()
            flags = config.get("automation", {}).get("notify_flags", {})

            # 兼容旧配置
            if not flags:
                mode = config.get("automation", {}).get("notify_mode", "popup")
                flags = {"in_app": False, "popup": mode == "popup", "tray": mode == "tray"}

            # 如果所有通知都关闭，直接返回
            if not any(flags.get(k) for k in ("in_app", "popup", "tray")):
                return

            # 1. 任务栏闪烁提醒 (只要任务完成，且窗口不在前台，就闪烁)
            cls._flash_taskbar(parent_window)

            # 2. 应用内提醒 (Toast) - 必须勾选才显示
            if flags.get("in_app"):
                cls._show_toast(title, message, parent_window)

            # 3. 桌面通知 (Windows 右下角) - 必须勾选才显示
            if flags.get("tray"):
                cls._show_desktop_notification(title, message, has_issue)

            # 4. 弹框通知 (自定义 QDialog) - 必须勾选才显示
            if flags.get("popup"):
                cls._show_popup_dialog(title, message, parent_window, on_view_callback)

        except Exception as e:
            log_error(f'[NotificationManager] 通知异常: {e}')

    @classmethod
    def _flash_taskbar(cls, window):
        """让任务栏图标闪烁"""
        if window and not window.isActiveWindow():
            QApplication.alert(window, 2000)  # 闪烁 2 秒
            # 如果需要持续闪烁，可以加个定时器，但通常 alert 足够引起注意

    @classmethod
    def _show_toast(cls, title, message, window):
        """应用内 Toast 提醒"""
        toast = QLabel(window)
        icon = "⚠️" if "异常" in title else "✅"
        toast.setText(f" {icon} {title} <br> {message} ")
        toast.setWordWrap(True)
        toast.setMaximumWidth(400)
        # 现代卡片样式
        toast.setStyleSheet(
            "background: rgba(30, 41, 59, 0.95); color: #f8fafc;"
            "border: 1px solid #3b82f6; border-radius: 8px;"
            "padding: 12px 16px; font-size: 13px; font-weight: 500;"
            "box-shadow: 0 4px 6px rgba(0,0,0,0.1);"
        )
        toast.adjustSize()
        # 显示在窗口右上角
        toast.move(max(10, window.width() - toast.width() - 20), 80)
        toast.show()
        toast.raise_()
        # 4秒后自动消失
        QTimer.singleShot(4000, toast.deleteLater)

    @classmethod
    def _show_desktop_notification(cls, title, message, has_issue):
        """Windows 桌面右下角通知"""
        try:
            # 加载应用图标 (解决显示 Python 图标的问题)
            icon_path = get_resource_path("ui/logo.ico")
            if not os.path.exists(icon_path):
                icon_path = get_resource_path("ui/logo.png")

            tray_icon = QIcon(icon_path) if os.path.exists(icon_path) else None

            # 复用或创建 QSystemTrayIcon
            if cls._tray_instance is None:
                cls._tray_instance = QSystemTrayIcon(tray_icon)
                cls._tray_instance.setToolTip("COSMIC 智能评估")
                cls._tray_instance.show()  # 必须在系统托盘区显示图标，通知才能弹出来

            # 根据是否有异常选择通知类型
            msg_type = QSystemTrayIcon.Warning if has_issue else QSystemTrayIcon.Information
            cls._tray_instance.showMessage(title, message, msg_type, 6000)
        except Exception as e:
            log_error(f'[NotificationManager] 桌面通知失败: {e}')

    @classmethod
    def _show_popup_dialog(cls, title, message, window, on_view_callback):
        """自定义弹框通知"""
        # 关闭旧的弹框
        if hasattr(window, "_notify_dialog") and window._notify_dialog:
            try:
                window._notify_dialog.close()
            except:
                pass

        dlg = QDialog(window)
        dlg.setWindowTitle("任务完成提醒")
        dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowStaysOnTopHint)  # 置顶
        dlg.setModal(False)
        dlg.resize(420, 160)

        lay = QVBoxLayout(dlg)
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 16px; font-weight: 700; color: #2563eb;")
        lay.addWidget(title_label)

        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet("font-size: 13px; color: #475569; margin-top: 5px;")
        lay.addWidget(msg_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_view = QPushButton("查看结果")
        btn_view.setStyleSheet(
            "background: #2563eb; color: white; border-radius: 4px; padding: 5px 15px; font-weight: bold;")
        btn_close = QPushButton("关闭")
        btn_row.addWidget(btn_view)
        btn_row.addWidget(btn_close)
        lay.addLayout(btn_row)

        if on_view_callback:
            btn_view.clicked.connect(lambda: (dlg.close(), on_view_callback()))
        else:
            btn_view.clicked.connect(dlg.close)
        btn_close.clicked.connect(dlg.close)

        # 15秒自动关闭
        QTimer.singleShot(15000, dlg.close)

        window._notify_dialog = dlg
        dlg.show()