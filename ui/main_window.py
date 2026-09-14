import os
import random
from datetime import datetime

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QScrollArea,
    QApplication,
    QStackedWidget,
    QMessageBox, QDialog,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QPalette
from .task_card import TaskCard
from .upload_dialog import UploadDialog  # ✅ 使用相对导入
from .re_review_dialog import ReReviewUploadDialog  # ✅ 新增重评对话框
from .re_review_card import ReReviewTaskCard  # ✅ 新增重评卡片
from .receipt_dialog import ReceiptFileDialog
from .evaluation_dialog import EvaluationDialog  # ✅ 新增评估对话框
from .evaluation_card import EvaluationTaskCard  # ✅ 新增评估卡片
from .sidebar import Sidebar  # ✅ 导入侧边栏
from .settings_widget import SettingsWidget  # ✅ 导入设置组件
from extend.matcher_config import MatcherConfig  # ✅ 导入配置类
from utils.path_utils import get_resource_path, open_directory, clear_directory
from utils.re_review_processor import ReReviewWorker  # ✅ 导入 Worker
from utils.themes import get_theme_stylesheet, apply_dark_title_bar
from utils.runtime_logger import log_info


class CosmicMainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cosmic 智能审核队列")
        self.setMinimumSize(1200, 800)

        # 默认窗口尺寸：取可用屏幕的约 80%（上限 1560x960），保证侧边栏展开时
        # 卡片内容也放得下；小屏自动回退到最小尺寸
        screen = QApplication.primaryScreen()
        if screen is not None:
            avail = screen.availableGeometry()
            self.resize(
                max(1200, min(1560, int(avail.width() * 0.8))),
                max(800, min(960, int(avail.height() * 0.85))),
            )

        # 🚀 优先从配置加载主题状态，否则自适应系统
        config = MatcherConfig.load()
        self.is_dark_mode = config.get("theme", {}).get("is_dark", False)

        # 设置窗口图标
        # 如果你有 logo.png 请放在 ui 目录下，这里先写逻辑
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.ico")))

        # 主容器
        self.central_widget = QWidget()
        self.central_widget.setObjectName("CentralWidget")
        self.setCentralWidget(self.central_widget)

        # 全局水平布局 (左侧边栏 + 右侧主内容)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # 1. 侧边栏
        self.sidebar = Sidebar()
        self.sidebar.nav_changed.connect(self.switch_page)
        self.sidebar.theme_toggled.connect(self.toggle_theme)
        self.main_layout.addWidget(self.sidebar)

        # 2. 右侧垂直容器
        self.right_container = QWidget()
        self.right_container.setObjectName("RightContainer")
        self.right_layout = QVBoxLayout(self.right_container)
        self.right_layout.setContentsMargins(0, 0, 0, 0)
        self.right_layout.setSpacing(0)
        self.main_layout.addWidget(self.right_container)

        # 3. 栈容器 (存放不同页面)
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.setObjectName("StackedWidget")
        self.right_layout.addWidget(self.stacked_widget)

        # 创建页面
        self.page_home = self._create_home_page()
        self.page_initial_review = self._create_initial_review_page()
        self.page_receipt = self._create_receipt_page()
        self.page_re_review = self._create_re_review_page()  # ✅ 使用正式的重评页面
        # self.page_evaluation = self._create_evaluation_page()  # ✅ 新增评估页面
        self.page_settings = SettingsWidget()  # ✅ 使用正式的设置页面
        from .history_page import HistoryPage  # ✅ [NEW] 历史项目页
        self.page_history = HistoryPage()
        # [NEW] 历史项目：重新初评 / 断点续跑
        self.page_history.re_run_requested.connect(self._history_re_run)
        self.page_history.resume_requested.connect(self._history_resume)
        
        self.stacked_widget.addWidget(self.page_home)
        self.stacked_widget.addWidget(self.page_initial_review)
        self.stacked_widget.addWidget(self.page_receipt)
        self.stacked_widget.addWidget(self.page_re_review)
        self.stacked_widget.addWidget(self.page_history)
        # self.stacked_widget.addWidget(self.page_evaluation)
        self.stacked_widget.addWidget(self.page_settings)

        # 监听设置更新
        self.page_settings.config_updated.connect(self._apply_theme)

        # 默认显示主页
        self.stacked_widget.setCurrentIndex(0)
        self.sidebar.items[0].setSelected(True)

        # 应用初始主题
        self._apply_theme()



    def switch_page(self, index):
        # [NEW] 进入历史项目页时自动刷新
        if getattr(self, "page_history", None) is not None and index == 4:
            self.stacked_widget.setCurrentIndex(index)
            # 【优化】延迟刷新：先让页面完成绘制再启动后台扫描，点击切换不卡顿；
            # 大任务运行期间 refresh 内部会自动降级为提示页
            QTimer.singleShot(60, self.page_history.refresh)
            return
        self.stacked_widget.setCurrentIndex(index)

    def _history_re_run(self, raw_info):
        """[NEW] 历史项目：打开新建审核任务对话框并预填上次的材料与选项，
        用户可调整任意选项后自行点击开始"""
        info = dict(raw_info or {})
        info.pop("resume_state", None)
        info.pop("resume_source_paused_path", None)
        self.sidebar.on_item_clicked(1)
        dialog = UploadDialog(self, prefill_info=info)
        dialog.task_submitted.connect(self.add_new_task_and_navigate)
        dialog.exec()

    def _history_resume(self, raw_info):
        """[NEW] 历史项目：从暂停进度继续执行"""
        info = dict(raw_info or {})
        if not info.get("resume_state"):
            return
        self.add_new_task_and_navigate(info)

    def _create_home_page(self):
        page = QWidget()
        page.setObjectName("HomePage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        # 欢迎语 - 改进设计
        welcome_frame = QFrame()
        welcome_frame.setObjectName("WelcomeFrame")
        welcome_layout = QVBoxLayout(welcome_frame)
        welcome_layout.setContentsMargins(35, 25, 35, 25)
        welcome_layout.setSpacing(10)

        welcome_label = QLabel("🚀 欢迎使用 Cosmic 智能审核系统")
        welcome_label.setObjectName("WelcomeLabel")
        welcome_label.setStyleSheet(
            "font-size: 28px; font-weight: bold; color: #2563eb;"
        )
        welcome_label.setProperty("class", "welcome-title")

        # 获取当前时间进行个性化问候
        current_hour = datetime.now().hour
        if current_hour < 12:
            greeting = "上午好!"
        elif current_hour < 18:
            greeting = "下午好！"
        else:
            greeting = "晚上好！"

        greeting_label = QLabel(greeting)
        greeting_label.setStyleSheet("font-size: 14px; color: #64748b;")
        greeting_label.setProperty("class", "greeting-text")

        welcome_layout.addWidget(welcome_label)
        welcome_layout.addWidget(greeting_label)
        layout.addWidget(welcome_frame)

        # 快捷入口区域
        grid_layout = QHBoxLayout()
        grid_layout.setSpacing(30)

        # 存放快捷按钮以便更新主题
        self.shortcut_btns = []

        # 1. 业务操作卡片
        biz_group = QFrame()
        biz_group.setObjectName("BizGroup")
        biz_layout = QVBoxLayout(biz_group)
        biz_layout.setContentsMargins(30, 30, 30, 30)
        biz_layout.setSpacing(16)

        biz_title = QLabel("🚀 业务操作")
        biz_title.setObjectName("SectionTitle")
        biz_layout.addWidget(biz_title)

        upload_btn = self._create_shortcut_btn(
            "新增审核任务", "#2563eb", self.show_upload_dialog
        )
        receipt_btn = self._create_shortcut_btn(
            "新增评估确认单", "#10b981", self.show_receipt_dialog
        )
        re_review_btn = self._create_shortcut_btn(
            "新增重评任务", "#f59e0b", self.show_re_review_dialog
        )
        # evaluation_btn = self._create_shortcut_btn(
        #     "新增COSMIC评估", "#8b5cf6", self.show_evaluation_dialog
        # )

        biz_layout.addWidget(upload_btn)
        biz_layout.addWidget(receipt_btn)
        biz_layout.addWidget(re_review_btn)
        # biz_layout.addWidget(evaluation_btn)

        biz_layout.addSpacing(20)

        open_initial_btn = self._create_shortcut_btn(
            "打开初评文件夹", "#3b82f6", lambda: self.open_storage_dir("initial_review")
        )
        open_receipt_btn = self._create_shortcut_btn(
            "打开回单文件夹", "#10b981", lambda: self.open_storage_dir("receipt")
        )
        open_re_review_btn = self._create_shortcut_btn(
            "打开重评文件夹", "#f59e0b", lambda: self.open_storage_dir("re_review")
        )

        biz_layout.addWidget(open_initial_btn)
        biz_layout.addWidget(open_receipt_btn)
        biz_layout.addWidget(open_re_review_btn)

        biz_layout.addStretch()

        # 2. 系统清理卡片
        clean_group = QFrame()
        clean_group.setObjectName("CleanGroup")
        clean_layout = QVBoxLayout(clean_group)
        clean_layout.setContentsMargins(30, 30, 30, 30)
        clean_layout.setSpacing(16)

        clean_title = QLabel("🧹 系统清理")
        clean_title.setObjectName("SectionTitle")
        clean_layout.addWidget(clean_title)

        clean_initial = self._create_shortcut_btn(
            "清空初评文件", "#6366f1", lambda: self.clear_files("initial_review")
        )
        clean_logs = self._create_shortcut_btn(
            "清空日志文件", "#8b5cf6", lambda: self.clear_files("logs")
        )
        clean_receipt = self._create_shortcut_btn(
            "清空回单文件", "#ec4899", lambda: self.clear_files("receipt")
        )
        clean_re_review = self._create_shortcut_btn(
            "清空重评文件", "#f43f5e", lambda: self.clear_files("re_review")
        )

        clean_temp = self._create_shortcut_btn(
            "清空临时压缩文件", "#14b8a6", lambda: self.clear_files("extraction")
        )

        all_clear_btn = QPushButton("⚠️ 一键清理全部数据")
        all_clear_btn.setFixedHeight(50)
        all_clear_btn.setCursor(Qt.PointingHandCursor)
        all_clear_btn.setProperty("class", "ShortcutBtn")
        all_clear_btn.setProperty("theme_color", "#ef4444")
        if not hasattr(self, "shortcut_btns"):
            self.shortcut_btns = []
        self.shortcut_btns.append(all_clear_btn)
        self._refresh_shortcut_btn_style(all_clear_btn)
        all_clear_btn.clicked.connect(self.clear_all_files)

        clean_layout.addWidget(clean_initial)
        clean_layout.addWidget(clean_logs)
        clean_layout.addWidget(clean_receipt)
        clean_layout.addWidget(clean_re_review)
        clean_layout.addWidget(clean_temp)

        clean_layout.addSpacing(15)

        clean_layout.addWidget(all_clear_btn)
        clean_layout.addStretch()

        grid_layout.addWidget(biz_group, 1)

        # 添加分隔线（竖直，中间对齐）
        # separator = QLabel("|\n" * 20)
        # separator.setAlignment(Qt.AlignCenter)
        # separator.setProperty("class", "separator-line")
        # grid_layout.addWidget(separator, 0)

        grid_layout.addWidget(clean_group, 1)
        layout.addLayout(grid_layout)
        layout.addStretch()

        return page

    def _create_shortcut_btn(self, text, color, slot):
        btn = QPushButton(text)
        btn.setFixedHeight(45)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setProperty("class", "ShortcutBtn")

        # 保存原始颜色以供主题切换时使用
        btn.setProperty("theme_color", color)
        if not hasattr(self, "shortcut_btns"):
            self.shortcut_btns = []
        self.shortcut_btns.append(btn)

        self._refresh_shortcut_btn_style(btn)
        btn.clicked.connect(slot)
        return btn

    def _refresh_shortcut_btn_style(self, btn):
        color = btn.property("theme_color")
        is_red_btn = color == "#ef4444" or color == "#ee0000"

        if self.is_dark_mode:
            if is_red_btn:
                # 深色模式下的红色按钮
                btn.setStyleSheet(
                    f"""
                    QPushButton {{
                        color: #ffffff;
                        border: 1px solid rgba(239, 68, 68, 0.6);
                        background-color: #ef4444;
                        border-radius: 10px;
                        font-size: 15px;
                        font-weight: 600;
                        padding: 14px 16px;
                        text-align: center;
                        transition: all 0.3s ease;
                    }}
                    QPushButton:hover {{
                        background-color: #dc2626;
                        border: 1px solid #991b1b;
                    }}
                    QPushButton:pressed {{
                        background-color: #b91c1c;
                        transform: scale(0.98);
                    }}
                """
                )
            else:
                # 深色模式：半透明卡片效果，带颜色边框
                btn.setStyleSheet(
                    f"""
                    QPushButton {{
                        color: #e3f2fd;
                        border: 1px solid rgba(100, 200, 255, 0.4);
                        background-color: rgba(30, 41, 59, 0.5);
                        border-radius: 10px;
                        font-size: 15px;
                        font-weight: 600;
                        padding: 14px 16px;
                        text-align: center;
                        transition: all 0.3s ease;
                    }}
                    QPushButton:hover {{
                        background-color: rgba(51, 65, 85, 0.8);
                        border: 1px solid {color};
                        color: {color};
                    }}
                    QPushButton:pressed {{
                        background-color: rgba(30, 41, 59, 0.8);
                        transform: scale(0.98);
                    }}
                """
                )
        else:
            if is_red_btn:
                # 浅色模式下的红色按钮
                btn.setStyleSheet(
                    f"""
                    QPushButton {{
                        color: #ffffff;
                        border: 2px solid #ef4444;
                        background-color: #ef4444;
                        border-radius: 12px;
                        font-size: 16px;
                        font-weight: 600;
                        padding: 12px;
                        text-align: center;
                        transition: all 0.3s ease;
                    }}
                    QPushButton:hover {{
                        background-color: #dc2626;
                        border: 2px solid #dc2626;
                    }}
                    QPushButton:pressed {{
                        background-color: #b91c1c;
                        transform: scale(0.98);
                    }}
                """
                )
            else:
                # 浅色模式：现代卡片设计，添加阴影和过渡
                btn.setStyleSheet(
                    f"""
                    QPushButton {{
                        color: #1e293b;
                        border: 2px solid #e2e8f0;
                        background-color: #ffffff;
                        border-radius: 12px;
                        font-size: 16px;
                        font-weight: 600;
                        padding: 12px;
                        text-align: center;
                        transition: all 0.3s ease;
                    }}
                    QPushButton:hover {{
                        border-color: {color};
                        border-width: 2px;
                        background-color: #f8fafc;
                        color: {color};
                    }}
                    QPushButton:pressed {{
                        background-color: #f1f5f9;
                        transform: scale(0.98);
                    }}
                """
                )

    def clear_files(self, storage_key):
        config = MatcherConfig.load()
        path = config.get("storage", {}).get(storage_key)
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "警告", f"配置的路径不存在: {path}")
            return

        reply = QMessageBox.question(
            self,
            "确认清理",
            f"确定要清空该目录吗？\n{path}",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            from utils.path_utils import clear_directory

            success, msg = clear_directory(path)
            if success:
                QMessageBox.information(self, "完成", "清理完毕")
            else:
                QMessageBox.critical(self, "错误", msg)

    def clear_all_files(self):
        reply = QMessageBox.question(
            self,
            "危险操作",
            "确定要清空初评、日志、回单、重评以及临时压缩全部目录吗？",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        from utils.path_utils import clear_directory

        config = MatcherConfig.load()
        keys = ["initial_review", "logs", "receipt", "re_review", "extraction"]
        results = []
        for key in keys:
            path = config.get("storage", {}).get(key)
            if path and os.path.exists(path):
                clear_directory(path)
                results.append(f"{key}: 已清理")

        QMessageBox.information(self, "完成", "\n".join(results))

    def open_storage_dir(self, storage_key):
        config = MatcherConfig.load()
        path = config.get("storage", {}).get(storage_key)
        if path and os.path.exists(path):
            open_directory(path)
        else:
            QMessageBox.warning(self, "警告", f"目录不存在或未配置: {path}")

    def _create_placeholder_page(self, text):
        page = QWidget()
        layout = QVBoxLayout(page)
        label = QLabel(text)
        label.setStyleSheet("font-size: 24px;")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        return page

    def _create_initial_review_page(self):
        """创建初评页面 (即原有的主界面内容)"""
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 20)
        page_layout.setSpacing(20)

        # Header (仍保留在页面内，或作为全局 Header)
        self.header = self._create_header()
        page_layout.addWidget(self.header)

        # 任务列表 (滚动区域)
        scroll_container = QWidget()
        scroll_layout = QVBoxLayout(scroll_container)
        scroll_layout.setContentsMargins(20, 0, 20, 0)
        scroll_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        # [UI] 纵向滚动条细化并融入深浅色背景；横向兜底：卡片步骤条在极窄窗口下
        # 允许出横向滚动条而不是把右侧按钮裁掉（执行过程中的自适应）
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet(
            """
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical { background: transparent; width: 8px; margin: 2px; }
            QScrollBar::handle:vertical { background: rgba(100,116,139,0.5); border-radius: 4px; min-height: 30px; }
            QScrollBar::handle:vertical:hover { background: rgba(100,116,139,0.8); }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; width: 0px; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
            QScrollBar:horizontal { background: transparent; height: 8px; margin: 2px; }
            QScrollBar::handle:horizontal { background: rgba(100,116,139,0.5); border-radius: 4px; min-width: 30px; }
            QScrollBar::handle:horizontal:hover { background: rgba(100,116,139,0.8); }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { height: 0px; width: 0px; }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }
            """
        )
        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        self.task_layout = QVBoxLayout(scroll_content)
        self.task_layout.setSpacing(20)

        # 初始不再加载模拟数据
        self.task_layout.addStretch()
        scroll.setWidget(scroll_content)
        scroll_layout.addWidget(scroll)
        page_layout.addWidget(scroll_container)

        return page

    def _apply_theme(self):
        """应用主题样式"""
        # 从配置中同步主题状态
        config = MatcherConfig.load()
        self.is_dark_mode = config.get("theme", {}).get("is_dark", False)

        theme = get_theme_stylesheet(self.is_dark_mode)
        QApplication.instance().setStyleSheet(theme)

        # 重新 polish 侧边栏项目以应用新的样式
        for item in self.sidebar.items:
            item.style().unpolish(item)
            item.style().polish(item)
            # 也更新子元素
            item.icon_label.style().unpolish(item.icon_label)
            item.icon_label.style().polish(item.icon_label)
            item.text_label.style().unpolish(item.text_label)
            item.text_label.style().polish(item.text_label)

        # 刷新主页快捷按钮样式 (因为它们含有内联样式)
        if hasattr(self, "shortcut_btns"):
            for btn in self.shortcut_btns:
                self._refresh_shortcut_btn_style(btn)

        # 刷新初评页面的任务卡片
        if hasattr(self, "task_layout"):
            for i in range(self.task_layout.count()):
                item = self.task_layout.itemAt(i)
                if item and item.widget() and isinstance(item.widget(), TaskCard):
                    item.widget().update_style()

        # 刷新所有重评和回单卡片的样式
        if hasattr(self, "re_review_task_layout"):
            for i in range(self.re_review_task_layout.count()):
                item = self.re_review_task_layout.itemAt(i)
                if (
                        item
                        and item.widget()
                        and isinstance(item.widget(), ReReviewTaskCard)
                ):
                    item.widget().update_theme_style()

        if hasattr(self, "receipt_task_layout"):
            for i in range(self.receipt_task_layout.count()):
                item = self.receipt_task_layout.itemAt(i)
                if (
                        item
                        and item.widget()
                        and isinstance(item.widget(), ReReviewTaskCard)
                ):
                    item.widget().update_theme_style()

        # 历史项目页：卡片样式是渲染时的快照，主题切换后必须整页重绘，
        # 否则会出现深浅色卡片混杂
        if hasattr(self, "page_history") and hasattr(self.page_history, "apply_theme"):
            self.page_history.apply_theme()

        # 设置页：区块标题/描述/清空等快照样式重刷
        if hasattr(self, "page_settings") and hasattr(self.page_settings, "apply_theme"):
            self.page_settings.apply_theme()

        # 应用原生标题栏深色模式
        apply_dark_title_bar(self, self.is_dark_mode)

        # 更新主题按钮的样式和文字
        if hasattr(self, "sidebar") and hasattr(self.sidebar, "theme_btn"):
            self.sidebar._update_theme_btn_style()
            if self.is_dark_mode:
                self.sidebar.theme_btn.setText("☀️")
            else:
                self.sidebar.theme_btn.setText("🌙")

        if self.is_dark_mode:
            if hasattr(self, "theme_btn"):
                self.theme_btn.setText("☀️ 白天模式")
        else:
            if hasattr(self, "theme_btn"):
                self.theme_btn.setText("🌙 深色模式")

    def toggle_theme(self):
        """切换主题"""
        self.is_dark_mode = not self.is_dark_mode
        # 同步到配置
        config = MatcherConfig.load()
        if "theme" not in config:
            config["theme"] = {}
        config["theme"]["is_dark"] = self.is_dark_mode
        MatcherConfig.save(config)

        self._apply_theme()

        # 刷新所有重评卡片的样式
        for i in range(self.re_review_task_layout.count()):
            item = self.re_review_task_layout.itemAt(i)
            if item and item.widget() and isinstance(item.widget(), ReReviewTaskCard):
                item.widget().update_theme_style()

    def _create_header(self):
        """创建顶部 Header"""
        header = QFrame()
        header.setObjectName("HeaderFrame")
        header.setFixedHeight(80)  # 增加高度
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)  # 左右留白让内容居中感更好

        # 左侧信息
        left_info = QVBoxLayout()
        left_info.setSpacing(4)
        left_info.setAlignment(Qt.AlignVCenter)  # 垂直居中

        title = QLabel("🚀 Cosmic 智能审核队列")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        left_info.addWidget(title)

        header_layout.addLayout(left_info)
        header_layout.addStretch()

        # 右侧按钮组
        btn_group = QHBoxLayout()
        btn_group.setSpacing(10)
        btn_group.setAlignment(Qt.AlignVCenter)  # 垂直居中

        # 主题切换按钮
        self.theme_btn = QPushButton("🌙 深色模式")
        self.theme_btn.setObjectName("ThemeBtn")
        self.theme_btn.clicked.connect(self.toggle_theme)
        btn_group.addWidget(self.theme_btn)

        # 【新增】打开初评文件夹按钮
        open_dir_btn = QPushButton("📁 打开初评文件夹")
        open_dir_btn.setObjectName("ThemeBtn")  # 复用辅助按钮样式
        open_dir_btn.clicked.connect(self.open_initial_review_dir)
        btn_group.addWidget(open_dir_btn)

        upload_btn = QPushButton("新建审核任务")
        upload_btn.setObjectName("UploadBtn")
        # 移除硬编码样式
        upload_btn.clicked.connect(self.show_upload_dialog)
        btn_group.addWidget(upload_btn)

        header_layout.addLayout(btn_group)

        return header

    def open_initial_review_dir(self):
        """打开初评文件夹"""
        config = MatcherConfig.load()
        path = config.get("storage", {}).get("initial_review")
        if path:
            # 确保目录存在
            if not os.path.exists(path):
                try:
                    os.makedirs(path)
                except:
                    pass
            open_directory(path)
        else:
            QMessageBox.warning(self, "提示", "初评文件夹尚未配置。")

    def show_upload_dialog(self):
        dialog = UploadDialog(self)
        dialog.task_submitted.connect(self.add_new_task_and_navigate)
        dialog.exec()

    def add_new_task_and_navigate(self, task_info):
        """添加新任务并跳转到初评页面"""
        # 先执行跳转，让用户感知到页面切换
        self.sidebar.on_item_clicked(1)
        # 再异步或者立即添加任务
        self.add_new_task(task_info)
        # 强制界面刷新
        from PySide6.QtWidgets import QApplication

        QApplication.processEvents()

    def add_new_task(self, task_info):
        """动态添加新任务卡片，随机取色作为边框"""
        # 确定类型
        type_label = "结算"

        # 随机取色 (选取一些好看的 UI 色系)
        colors = [
            "#2563eb",
            "#10b981",
            "#f59e0b",
            "#a855f7",
            "#ef4444",
            "#06b6d4",
            "#ec4899",
        ]
        border_color = random.choice(colors)

        # 日志（1~6）
        timestamp = datetime.now().strftime("%H:%M:%S")

        # 处理初始日志，第一步现在交给 TaskCard 异步处理
        step1_status = "pending"
        step1_log = "等待比对任务开始..."
        if task_info.get("validation_results"):
            # 如果已经有结果（比如重载），则保持原有逻辑
            step1_status = "done"  # 示例简化

        # ✅ 10 个步骤，加入环境准备节点 0 和数据属性重复检测节点 9
        # 标签尽量 4-6 字，避免换行撑高步骤条
        steps = [
            ("环境准备", "pending"),
            ("模板校验", "pending"),
            ("空值检查", "pending"),
            ("送审比例", "pending"),
            ("附加值因子", "pending"),
            ("层级匹配", "pending"),
            ("功能过程", "pending"),
            ("数据移动类型", "pending"),
            ("资产清单", "pending"),
            ("属性查重", "pending"),
        ]

        logs = {
            0: "正在准备环境...",
            1: "等待执行...",
            2: "等待执行...",
            3: "等待执行...",
            4: "等待执行...",
            5: "等待执行...",
            6: "等待执行...",
            7: "等待执行...",
            8: "等待执行...",
            9: "等待执行...",
        }

        # 默认日志（显示在卡片底部）
        default_log = f"✅ 任务于 {timestamp} 提交，正在处理..."

        # 构造 task_data，完全匹配 mock_data 结构
        task_data = {
            "filename": task_info["filename"],
            "type_label": type_label,
            "days": task_info["days"],  # 字符串也可，UI 只做拼接
            "border_color": border_color,
            "steps": steps,  # ✅ 6 个元组，名称/状态对齐 mock
            "default_log": default_log,
            "logs": logs,  # ✅ 1~6 键完整
            "report_sections": [],  # 新任务暂无报告，但字段存在（兼容 ReportDialog）
            "raw_task_info": task_info,  # 保留原始数据
            # [NEW] 断点续跑：历史页"继续执行"传来的 resume_state 提到 task_data 顶层，
            #   TaskCard 可直接读取以从 completed_through 步继续
            #   （raw_task_info["resume_state"] 原位置仍保留，两处等价）
            "resume_state": task_info.get("resume_state"),
        }

        # 创建卡片并插入顶部
        card = TaskCard(task_data)
        self.task_layout.insertWidget(0, card)

    def _create_re_review_page(self):
        """创建重评页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header
        header = QFrame()
        header.setObjectName("HeaderFrame")
        header.setFixedHeight(80)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)

        title = QLabel("🔄 重评页面")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        self.re_review_btn = QPushButton("新建重评任务")
        self.re_review_btn.setObjectName("NewTaskBtn")

        self.re_review_btn.clicked.connect(self.show_re_review_dialog)
        header_layout.addWidget(self.re_review_btn)

        layout.addWidget(header)

        # Content - Scroll Area for Task Cards
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")

        scroll_content = QWidget()
        scroll_content.setObjectName("ReReviewScrollContent")
        self.re_review_task_layout = QVBoxLayout(scroll_content)
        self.re_review_task_layout.setContentsMargins(30, 20, 30, 20)
        self.re_review_task_layout.setSpacing(15)
        self.re_review_task_layout.addStretch()

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, stretch=1)

        return page

    def show_re_review_dialog(self):
        """显示重评弹窗"""
        dialog = ReReviewUploadDialog(self)
        dialog.task_started.connect(self.add_re_review_task)
        dialog.exec()

    def add_re_review_task(self, task_info):
        """启动重评任务并在页面显示进度卡片"""
        # 1. 自动跳转到重评页面 (Index 3)
        self.sidebar.on_item_clicked(3)

        # 2. 创建并显示卡片
        card = ReReviewTaskCard(task_info)
        self.re_review_task_layout.insertWidget(0, card)

        # 3. 启动后台线程
        project_name = os.path.basename(task_info["excel1"])
        worker = ReReviewWorker(
            task_info["excel1"],
            task_info["excel2"],
            task_info["output_dir"],
            project_name,
        )

        # 保持引用防止被垃圾回收
        # 【崩溃修复】严禁在自定义 finished/error 信号回调里把 worker 移出列表：
        # 这两个信号在 run() 内部 emit，此时线程可能仍在执行收尾代码，
        # 移除最后一个引用会销毁运行中的 QThread → Qt qFatal 闪退(0xc0000409)。
        # 改为：仅清理【已真正结束】(isRunning()=False) 的旧 worker。
        if not hasattr(self, "re_review_workers"):
            self.re_review_workers = []
        self.re_review_workers = [
            w for w in self.re_review_workers if w.isRunning()
        ]
        self.re_review_workers.append(worker)

        worker.progress.connect(card.update_progress)
        worker.finished.connect(lambda paths: card.set_completed(paths[0], paths[1]))
        worker.error.connect(card.set_error)

        # [NEW] 经队列管理器按并发配置启动
        from utils.task_queue import TaskQueueManager

        TaskQueueManager().submit("re_review", worker)

    def _create_receipt_page(self):
        """创建回单页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header
        header = QFrame()
        header.setObjectName("HeaderFrame")
        header.setFixedHeight(80)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)

        title = QLabel("📄 回单页面")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        self.receipt_btn = QPushButton("新增确认单")
        self.receipt_btn.setObjectName("ReceiptBtn")
        self.receipt_btn.clicked.connect(self.show_receipt_dialog)
        header_layout.addWidget(self.receipt_btn)

        layout.addWidget(header)

        # Content - Scroll Area for Task Cards
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")

        scroll_content = QWidget()
        scroll_content.setObjectName("ReceiptScrollContent")
        self.receipt_task_layout = QVBoxLayout(scroll_content)
        self.receipt_task_layout.setContentsMargins(30, 20, 30, 20)
        self.receipt_task_layout.setSpacing(15)
        self.receipt_task_layout.addStretch()

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, stretch=1)

        return page

    def show_receipt_dialog(self):
        """显示回单配置弹窗"""
        dialog = ReceiptFileDialog(self)
        dialog.generation_requested.connect(self.handle_receipt_generation)
        dialog.exec()

    def handle_receipt_generation(self, data):
        """处理回单生成逻辑"""
        # 1. 自动跳转到回单页面 (index 2)
        self.sidebar.on_item_clicked(2)

        # 检查是否为批量任务
        is_batch = data.get("is_batch", False)
        batch_tasks = data.get("tasks", []) if is_batch else [data]

        # 存储卡片以便后续更新 {index: card_widget}
        card_map = {}

        # 2. 为每个任务创建卡片
        # 注意：列表是倒序插入，为了保持顺序一致性，我们先生成卡片对象列表，再倒序插入布局
        new_cards = []
        for i, task_data in enumerate(batch_tasks):
            task_info = {
                "project_name": task_data["project_name"],
                "type": "回单生成",
                "time": datetime.now().strftime("%H:%M:%S"),
                "status_text": "等待处理..." if is_batch and i > 0 else "正在处理...",
            }
            card = ReReviewTaskCard(task_info, card_type="receipt")
            card.excel1_btn.setText("📊 评估报告(已回写)")
            card.excel2_btn.setText("📋 评估确认单(Word)")
            new_cards.append(card)
            card_map[i] = card

        # 倒序插入布局，这样第一个任务在最上面
        for card in reversed(new_cards):
            self.receipt_task_layout.insertWidget(0, card)

        from utils.receipt_processor import ReceiptWorker

        # 确定项目名称前缀
        if is_batch:
            project_base = data.get("project_name", "Batch_Receipts")
        elif data.get("is_merge"):
            groups = data.get("file_groups", {})
            first_group = list(groups.values())[0] if groups else {}
            report_for_name = first_group.get("eval_report", "Merged")
            project_base = os.path.basename(report_for_name)
        else:
            project_base = os.path.basename(data.get("eval_report_path", "Receipt"))

        # 3. 初始化 Worker (传入整个 data，如果是 batch，worker 会自己处理)
        worker = ReceiptWorker(data, project_base)

        if not hasattr(self, "receipt_workers"):
            self.receipt_workers = []
        self.receipt_workers.append(worker)

        # 4. 连接信号

        def on_item_finished(idx, result):
            if idx in card_map:
                card = card_map[idx]

                # Check for error first
                if result.get("error"):
                    card.set_error(result["error"])
                    return

                word_path = result.get("output_path")
                stats = result.get("stats", {})
                excel_reports = result.get("excel_reports", [])

                # 如果有多个 Excel 报告（如合并模式），传入列表以便 UI 适配显示为文件夹
                if len(excel_reports) > 1:
                    report_for_card = excel_reports
                else:
                    report_for_card = excel_reports[0] if excel_reports else None

                card.set_completed(excel1=report_for_card, excel2=word_path)
                card.log_label.setText(f"✅ 生成成功！已保存")
                card.task_info["output_dir"] = os.path.dirname(word_path)
                if stats:
                    card.update_stats(stats)
                    # 更新真实项目名称
                    if isinstance(stats, list):
                        # 如果是合并任务，stats 是一个列表，取最后一个（汇总）或尝试从列表中找
                        if stats:
                            summary = stats[-1]
                            if summary.get("real_project_name"):
                                card.update_title(summary["real_project_name"])
                    elif isinstance(stats, dict) and stats.get("real_project_name"):
                        card.update_title(stats["real_project_name"])

        worker.item_finished.connect(on_item_finished)

        # 处理整体完成
        def on_finished(result):
            # 如果不是批量模式，这里还需处理单个结果(兼容)
            if not is_batch:
                # 单个任务模式下，Worker 只会发 finished 不发 item_finished
                if 0 in card_map:  # 只有一个卡片
                    on_item_finished(0, result)

            # 批量模式全部完成后，可以做一些清理或通知
            # 自动化：打开文件夹 (取最后一个路径)
            if isinstance(result, dict) and result.get("output_path"):
                word_path = result.get("output_path")
                config = MatcherConfig.load()
                if config.get("automation", {}).get("auto_open", True):
                    open_directory(os.path.dirname(word_path))
            elif is_batch and batch_tasks:
                # 批量模式结束，打开第一个任务的目录即可
                output_dir = MatcherConfig.load().get("storage", {}).get("receipt")
                if output_dir and os.path.exists(output_dir):
                    if (
                            MatcherConfig.load()
                                    .get("automation", {})
                                    .get("auto_open", True)
                    ):
                        open_directory(output_dir)

        worker.finished.connect(on_finished)

        def on_error(msg):
            # 这里简单处理：如果是全局错误，把所有未完成的卡片都设为错误
            # 但实际上 Worker 内部处理每个 item 异常，不会轻易抛出全局异常
            # 除非是完全无法启动
            for card in card_map.values():
                # 只有还没完成的显示错误
                if card.files_widget.isHidden():
                    card.set_error(msg)

        worker.error.connect(on_error)
        # 【崩溃修复】不在信号回调里移除 worker（信号在 run() 内部发出，移除最后一个
        # 引用会销毁运行中的 QThread → qFatal 闪退）；改为创建新 worker 时清理已结束的
        if not hasattr(self, "receipt_workers"):
            self.receipt_workers = []
        self.receipt_workers = [w for w in self.receipt_workers if w.isRunning()]
        self.receipt_workers.append(worker)

        # [NEW] 经队列管理器按并发配置启动
        from utils.task_queue import TaskQueueManager

        TaskQueueManager().submit("receipt", worker)

    def _create_evaluation_page(self):
        """创建评估页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header
        header = QFrame()
        header.setObjectName("HeaderFrame")
        header.setFixedHeight(80)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)

        title = QLabel("📊 COSMIC 评估")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        # 新增评估按钮
        self.evaluation_btn = QPushButton("新增评估任务")
        self.evaluation_btn.setObjectName("EvaluationBtn")
        self.evaluation_btn.setFixedSize(150, 40)
        self.evaluation_btn.clicked.connect(self.show_evaluation_dialog)
        header_layout.addWidget(self.evaluation_btn)

        layout.addWidget(header)

        # Content - Scroll Area for Task Cards
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")

        scroll_content = QWidget()
        scroll_content.setObjectName("EvaluationScrollContent")
        self.evaluation_task_layout = QVBoxLayout(scroll_content)
        self.evaluation_task_layout.setContentsMargins(30, 20, 30, 20)
        self.evaluation_task_layout.setSpacing(15)
        self.evaluation_task_layout.addStretch()

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, stretch=1)

        return page

    def show_evaluation_dialog(self):
        """显示评估对话框"""
        dialog = EvaluationDialog(self)
        dialog.evaluation_requested.connect(self.start_evaluation_task)
        dialog.exec()

    def start_evaluation_task(self, task_info):
        """启动评估任务并显示卡片"""
        # 1. 添加卡片并跳转
        self.add_evaluation_task(task_info)

        # 2. 启动后台 Worker
        from utils.evaluation_processor import EvaluationWorker

        # 查找刚才添加的卡片以更新状态
        card = None
        for i in range(self.evaluation_task_layout.count()):
            item = self.evaluation_task_layout.itemAt(i)
            if item and item.widget() and isinstance(item.widget(), EvaluationTaskCard):
                if item.widget().task_id == task_info["task_id"]:
                    card = item.widget()
                    break

        if not card:
            return

        worker = EvaluationWorker(
            architecture_doc_path=task_info.get("architecture_file_path"),
            evaluation_excel_path=task_info["evaluation_excel_path"],
            architecture_text=task_info.get("architecture_input"),
            manday=task_info["manday"],
            is_manual_mode=True,
            is_auto_evaluate=task_info.get("is_auto_evaluate", False),
            project_name=task_info.get("project_name")
        )

        if not hasattr(self, "evaluation_workers"):
            self.evaluation_workers = []
        # 【崩溃修复】不在信号回调里移除 worker（同 re_review），改为清理已结束的
        self.evaluation_workers = [
            w for w in self.evaluation_workers if w.isRunning()
        ]
        self.evaluation_workers.append(worker)

        # 连接 Worker 信号
        def on_finished(output_path):
            card.update_status("已完成")
            # 记录当前日志路径到卡片（按项目查询项目日志）
            from utils.runtime_logger import RuntimeLogger
            log_path = RuntimeLogger.get_current_log_path(task_info.get("project_name"))
            card.log_path = log_path

            card.set_output_file(output_path)
            # 在卡片上更新
            card.view_btn.setVisible(True)
            card.log_btn.setVisible(True)

        def on_error(msg):
            card.update_status("失败")
            QMessageBox.critical(self, "评估出错", f"项目 {task_info['project_name']} 评估失败:\n{msg}")

        # 新增项目：处理人工验证请求
        def on_verify_requested(results, callback):
            from ui.evaluation_dialog import ResultVerificationDialog
            dialog = ResultVerificationDialog(results, self, source_excel=worker.evaluation_excel_path)
            if dialog.exec() == QDialog.Accepted:
                callback(dialog.final_results)
            else:
                # 如果用户取消，默认传回原结果或空
                callback(results)

        worker.finished.connect(on_finished)
        worker.error.connect(on_error)
        worker.verify_requested.connect(on_verify_requested)
        worker.start()

        card.update_status("评估中...")

    def add_evaluation_task(self, task_info):
        """添加评估任务卡片"""
        # 自动跳转到评估页面 (Index 4)
        self.sidebar.on_item_clicked(4)

        # 创建并显示卡片
        card = EvaluationTaskCard(
            task_id=task_info.get("task_id", ""),
            project_name=task_info.get("project_name", "评估任务"),
            architecture_file=task_info.get("architecture_file", ""),
            evaluation_file=task_info.get("evaluation_excel_path") or task_info.get("evaluation_file", ""),
            output_file=task_info.get("output_file"),
            log_path=task_info.get("log_path"),
            status=task_info.get("status", "待评估"),
        )

        # 连接信号
        card.view_clicked.connect(self.open_file_location)
        card.log_clicked.connect(lambda: self.open_file_location(card.log_path))
        card.delete_clicked.connect(self.delete_evaluation_task)

        self.evaluation_task_layout.insertWidget(0, card)

    def open_file_location(self, file_path):
        """打开文件 (自动关联程序)"""
        if file_path and os.path.exists(file_path):
            open_directory(file_path)
        else:
            QMessageBox.warning(self, "错误", "文件不存在")

    def delete_evaluation_task(self, task_id):
        """删除评估任务卡片"""
        reply = QMessageBox.question(
            self,
            "确认删除",
            "确定要删除此评估任务吗？",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            # 从布局中移除卡片
            for i in range(self.evaluation_task_layout.count()):
                item = self.evaluation_task_layout.itemAt(i)
                if (
                        item
                        and item.widget()
                        and isinstance(item.widget(), EvaluationTaskCard)
                ):
                    if item.widget().task_id == task_id:
                        widget = item.widget()
                        self.evaluation_task_layout.removeWidget(widget)
                        widget.deleteLater()
                        break

    def closeEvent(self, event):
        """[NEW] 拦截窗口关闭，优雅停止后台任务并保存进度"""
        log_info('[CloseEvent] 检测到窗口关闭，正在保存任务进度...')

        # 1. 停止所有正在运行的初评任务
        if hasattr(self, "task_layout"):
            for i in range(self.task_layout.count()):
                item = self.task_layout.itemAt(i)
                if item and item.widget() and hasattr(item.widget(), "worker"):
                    card = item.widget()
                    if card.worker and card.worker.isRunning():
                        log_info(f"[CloseEvent] 正在停止任务: {card.task_data.get('filename')}")
                        # 触发停止标志
                        card.worker.stop()
                        card.worker._is_running = False
                        card.is_running = False

                        # [核心] 等待线程结束，给 finally 块时间执行 _dump_pause_state
                        # 最多等待 5 秒，防止程序卡死
                        card.worker.wait(5000)

                        # 2. 接受关闭事件
        event.accept()
        log_info('[CloseEvent] 进度保存完毕，程序关闭。')