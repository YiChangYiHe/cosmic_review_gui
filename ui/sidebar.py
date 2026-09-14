from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFrame,
    QSpacerItem,
    QSizePolicy,
)
from PySide6.QtCore import Qt, Signal, QPropertyAnimation, QEasingCurve, QSize
from PySide6.QtGui import QIcon, QFont


class SidebarItem(QFrame):
    """侧边栏单元项目"""

    clicked = Signal()

    def __init__(self, icon_text, label_text, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("SidebarItem")
        self.setFixedHeight(50)

        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(15, 0, 15, 0)
        self.layout.setSpacing(15)

        # 图标 (支持文字或图片)
        self.icon_label = QLabel(icon_text)
        self.icon_label.setFixedSize(28, 28)
        self.icon_label.setAlignment(Qt.AlignCenter)
        self.icon_label.setObjectName("SidebarItemIcon")
        self.icon_label.setScaledContents(False)  # 禁用缩放，避免emoji图标被压缩

        # 文字
        self.text_label = QLabel(label_text)
        self.text_label.setObjectName("SidebarItemText")

        self.layout.addWidget(self.icon_label)
        self.layout.addWidget(self.text_label)

        self.is_selected = False

    def setSelected(self, selected):
        self.is_selected = selected
        if selected:
            self.setProperty("selected", True)
        else:
            self.setProperty("selected", False)
        self.style().unpolish(self)
        self.style().polish(self)
        # 强制更新子元素样式
        self.icon_label.style().unpolish(self.icon_label)
        self.icon_label.style().polish(self.icon_label)
        self.text_label.style().unpolish(self.text_label)
        self.text_label.style().polish(self.text_label)

    def setExpanded(self, expanded):
        self.text_label.setVisible(expanded)
        if expanded:
            self.layout.setSpacing(15)
            self.layout.setContentsMargins(15, 0, 15, 0)
            self.layout.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        else:
            self.layout.setSpacing(0)
            self.layout.setContentsMargins(0, 0, 0, 0)
            self.layout.setAlignment(Qt.AlignCenter)

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class Sidebar(QFrame):
    """可折叠侧边栏"""

    nav_changed = Signal(int)  # 发送页面索引切换信号
    theme_toggled = Signal()  # 发送主题切换信号

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self.setFixedWidth(200)
        self.is_expanded = True

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 10, 0, 10)
        self.layout.setSpacing(5)

        # 顶部 切换按钮
        self.toggle_btn = QPushButton("☰")
        self.toggle_btn.setFixedSize(40, 40)
        self.toggle_btn.setCursor(Qt.PointingHandCursor)
        self.toggle_btn.setObjectName("SidebarToggleBtn")
        self.toggle_btn.clicked.connect(self.toggle_sidebar)

        self.toggle_layout = QHBoxLayout()
        self.toggle_layout.setContentsMargins(15, 0, 15, 10)
        self.toggle_layout.setAlignment(Qt.AlignLeft)
        self.toggle_layout.addWidget(self.toggle_btn)
        self.layout.addLayout(self.toggle_layout)

        # 导航项
        self.items = []
        self.nav_data = [
            ("🏠", "Cosmic 主页"),
            ("📝", "初评"),
            ("📄", "回单"),
            ("🔄", "重评"),
            ("🕘", "历史项目"),
            # ("📊", "评估"),
            ("⚙️", "设置"),
        ]

        for i, (icon, label) in enumerate(self.nav_data):
            item = SidebarItem(icon, label)
            item.clicked.connect(lambda idx=i: self.on_item_clicked(idx))
            self.layout.addWidget(item)
            self.items.append(item)

        self.layout.addStretch()

        # 底部 切换主题按钮 - 改进设计
        self.theme_btn = QPushButton("🌙")
        self.theme_btn.setFixedSize(48, 48)
        self.theme_btn.setCursor(Qt.PointingHandCursor)
        self.theme_btn.setObjectName("SidebarThemeBtn")
        self.theme_btn.setToolTip("点击切换主题 (Light/Dark)")
        self.theme_btn.clicked.connect(self.theme_toggled.emit)

        # 设置主题按钮样式
        self._update_theme_btn_style()

        self.theme_layout = QHBoxLayout()
        self.theme_layout.setContentsMargins(15, 0, 15, 10)
        self.theme_layout.setAlignment(Qt.AlignLeft)
        self.theme_layout.addWidget(self.theme_btn)
        self.layout.addLayout(self.theme_layout)

    def on_item_clicked(self, index):
        # 处理选中状态
        for i, item in enumerate(self.items):
            item.setSelected(i == index)
        self.nav_changed.emit(index)

    def toggle_sidebar(self):
        target_width = 70 if self.is_expanded else 200

        self.animation = QPropertyAnimation(self, b"minimumWidth")
        self.animation.setDuration(300)
        self.animation.setStartValue(self.width())
        self.animation.setEndValue(target_width)
        self.animation.setEasingCurve(QEasingCurve.InOutQuart)

        self.animation2 = QPropertyAnimation(self, b"maximumWidth")
        self.animation2.setDuration(300)
        self.animation2.setStartValue(self.width())
        self.animation2.setEndValue(target_width)
        self.animation2.setEasingCurve(QEasingCurve.InOutQuart)

        self.is_expanded = not self.is_expanded

        # 处理切换按钮和主题按钮的布局
        if self.is_expanded:
            self.toggle_layout.setContentsMargins(15, 0, 15, 10)
            self.toggle_layout.setAlignment(Qt.AlignLeft)
            self.theme_layout.setContentsMargins(15, 0, 15, 10)
            self.theme_layout.setAlignment(Qt.AlignLeft)
        else:
            self.toggle_layout.setContentsMargins(0, 0, 0, 10)
            self.toggle_layout.setAlignment(Qt.AlignCenter)
            self.theme_layout.setContentsMargins(0, 0, 0, 10)
            self.theme_layout.setAlignment(Qt.AlignCenter)

        # 先隐藏文字或改变布局，避免动画过程中由于宽度不足导致文字换行或闪烁
        for item in self.items:
            item.setExpanded(self.is_expanded)

        self.animation.start()
        self.animation2.start()

        # 切换按钮文字或图标
        self.toggle_btn.setText("☰" if self.is_expanded else "▶")

    def _update_theme_btn_style(self):
        """更新主题按钮样式，根据当前主题模式"""
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)

        if is_dark:
            # 深色模式：亮色按钮
            self.theme_btn.setStyleSheet(
                """
                QPushButton {
                    background-color: rgba(102, 204, 255, 0.1);
                    border: 2px solid #374151;
                    border-radius: 12px;
                    font-size: 24px;
                    color: #66ccff;
                    padding: 4px;
                    transition: all 0.3s ease;
                }
                QPushButton:hover {
                    background-color: rgba(102, 204, 255, 0.2);
                    border-color: #66ccff;
                    transform: scale(1.1);
                }
                QPushButton:pressed {
                    background-color: rgba(102, 204, 255, 0.15);
                    transform: scale(0.95);
                }
            """
            )
        else:
            # 浅色模式：深色按钮
            self.theme_btn.setStyleSheet(
                """
                QPushButton {
                    background-color: #f1f5f9;
                    border: 2px solid #e2e8f0;
                    border-radius: 12px;
                    font-size: 24px;
                    color: #0f172a;
                    padding: 4px;
                    transition: all 0.3s ease;
                }
                QPushButton:hover {
                    background-color: #dbeafe;
                    border-color: #2563eb;
                    transform: scale(1.1);
                }
                QPushButton:pressed {
                    background-color: #e0f2fe;
                    transform: scale(0.95);
                }
            """
            )
