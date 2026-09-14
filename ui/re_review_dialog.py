import os
import re
import json
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QMessageBox, QWidget, QGroupBox, QScrollArea,
)
from PySide6.QtCore import Qt, Signal, QByteArray, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QImage, QPixmap
from utils.path_utils import get_resource_path,clean_project_name


# 【新增】复用初评对话框的解析线程和下拉框组件
from .upload_dialog import UploadAreaWidget, ElidedLabel, DragOverlay, DownOnlyComboBox, ExcelParseWorker
from extend.matcher_config import MatcherConfig
from utils.archive_utils import ArchiveUtils
from utils.runtime_logger import log_error


class ReReviewFileItem(QFrame):
    """重评文件项"""
    remove_clicked = Signal(str)

    def __init__(self, file_path, is_eval_report, parent=None, is_dark=False):
        super().__init__(parent)
        self.file_path = file_path
        self.filename = os.path.basename(file_path)
        self.setProperty("class", "FileQueueItem")

        # 【深色模式适配】
        bg_color = "rgba(55, 65, 81, 0.5)" if is_dark else "rgba(243, 244, 246, 0.5)"
        text_color = "#f3f4f6" if is_dark else "#1f2937"

        self.setStyleSheet(f"""
            QFrame[class="FileQueueItem"] {{
                border-radius: 8px; padding: 10px; margin: 2px 0;
                background: {bg_color};
            }}
        """)
        layout = QHBoxLayout(self)
        icon_label = QLabel("X" if is_eval_report else "S")
        icon_label.setFixedSize(28, 28)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet(f"""
            background: {'#217346' if is_eval_report else '#2b579a'};
            color: white; border-radius: 4px; font-size: 11px; font-weight: bold;
        """)
        layout.addWidget(icon_label)
        name_label = ElidedLabel(self.filename)
        name_label.setStyleSheet(f"font-weight: 600; font-size: 12px; color: {text_color};")
        layout.addWidget(name_label, stretch=1)
        status_label = QLabel("旧评估报告" if is_eval_report else "新拆分表")
        status_label.setProperty("class", "task-meta")
        status_label.setFixedWidth(90)
        status_label.setAlignment(Qt.AlignCenter)
        status_label.setStyleSheet(f"""
            font-size: 11px; font-weight: 600;
            color: {'#10b981' if is_eval_report else '#3b82f6'};
            background: {'rgba(16, 185, 129, 0.15)' if is_eval_report else 'rgba(59, 130, 246, 0.15)'};
            border-radius: 4px; padding: 2px 6px;
        """)
        layout.addWidget(status_label)
        self.remove_btn = QPushButton()
        self.remove_btn.setFixedSize(28, 28)
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        trash_icon_svg = """
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
        </svg>
        """
        trash_qicon = QIcon(QPixmap.fromImage(QImage.fromData(QByteArray(trash_icon_svg.encode()))))
        self.remove_btn.setIcon(trash_qicon)
        self.remove_btn.setIconSize(QSize(16, 16))
        self.remove_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none; border-radius: 4px;
            }}
            QPushButton:hover {{
                background: rgba(239, 68, 68, 0.1);
            }}
        """)
        self.remove_btn.clicked.connect(lambda: self.remove_clicked.emit(self.file_path))
        layout.addWidget(self.remove_btn)


class ReReviewUploadDialog(QDialog):
    """重评文件上传对话框"""
    task_started = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新建重评任务")
        self.setMinimumSize(750, 700)
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.ico")))
        self.setAcceptDrops(True)

        from PySide6.QtWidgets import QApplication
        from utils.styles import apply_dark_title_bar
        self._is_dark = "background-color: #1f2937" in (QApplication.instance().styleSheet() or "")
        apply_dark_title_bar(self, self._is_dark)

        self.excel1_path = ""
        self.excel2_path = ""
        self.excel2_sheet_name = None
        self._parse_worker = None
        self.drag_overlay = DragOverlay(self)
        self.setup_ui()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drag_overlay.show_overlay()

    def dragLeaveEvent(self, event):
        self.drag_overlay.hide()

    def dropEvent(self, event: QDropEvent):
        self.drag_overlay.hide()
        if event.mimeData().hasUrls():
            files = [url.toLocalFile() for url in event.mimeData().urls()]
            self.handle_files(files)
            event.acceptProposedAction()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("border: none; background: transparent;")

        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(25)

        text_color = "#f3f4f6" if self._is_dark else "#1f2937"
        title = QLabel("新建重评任务")
        title.setStyleSheet(f"font-size: 24px; font-weight: bold; margin-bottom: 5px; color: {text_color};")
        layout.addWidget(title)

        file_section = QGroupBox("📂 文件选择 (支持拖拽 Excel)")
        border_color = "#38bdf8"
        file_section.setStyleSheet(f"""
            QGroupBox {{
                border: 2px solid {border_color}; border-radius: 12px; margin-top: 15px;
                padding-top: 15px; font-weight: bold; color: {border_color};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin; subcontrol-position: top left;
                left: 15px; padding: 0 5px;
            }}
        """)
        file_layout = QVBoxLayout(file_section)
        file_layout.setContentsMargins(20, 20, 20, 20)
        file_layout.setSpacing(15)

        self.upload_area = UploadAreaWidget()
        self.upload_area.files_dropped.connect(self.handle_files)
        file_layout.addWidget(self.upload_area)

        queue_label = QLabel("📋 已上传文件:")
        queue_label.setStyleSheet(f"font-size: 13px; font-weight: bold; margin-top: 5px; color: {text_color};")
        file_layout.addWidget(queue_label)

        self.files_scroll = QScrollArea()
        self.files_scroll.setWidgetResizable(True)
        self.files_scroll.setFixedHeight(150)
        bg_scroll = "rgba(31, 41, 55, 0.5)" if self._is_dark else "rgba(243, 244, 246, 0.5)"
        border_scroll = "#4b5563" if self._is_dark else "#d1d5db"
        self.files_scroll.setStyleSheet(f"""
            QScrollArea {{
                border: 1px solid {border_scroll}; border-radius: 8px;
                background-color: {bg_scroll};
            }}
        """)
        self.queue_content = QWidget()
        self.queue_content.setStyleSheet("background: transparent;")
        self.queue_layout = QVBoxLayout(self.queue_content)
        self.queue_layout.setSpacing(8)
        self.queue_layout.addStretch()
        self.files_scroll.setWidget(self.queue_content)
        file_layout.addWidget(self.files_scroll)

        # 【新增】拆分表工作表选择区域 (初始隐藏)
        self.sheet_selection_widget = QWidget()
        sheet_layout = QHBoxLayout(self.sheet_selection_widget)
        sheet_layout.setContentsMargins(0, 10, 0, 10)

        sheet_label = QLabel("📊 新拆分表工作表:")
        sheet_label.setStyleSheet(f"font-size: 13px; font-weight: 600; color: {text_color}; min-width: 120px;")

        self.sheet_combo = DownOnlyComboBox()
        self.sheet_combo.setFixedHeight(36)
        self.sheet_combo.setMaxVisibleItems(8)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)

        bg_input = "#374151" if self._is_dark else "#ffffff"
        text_input = "#f3f4f6" if self._is_dark else "#1f2937"
        border_input = "#4b5563" if self._is_dark else "#d1d5db"
        self.sheet_combo.setStyleSheet(f"""
            QComboBox {{
                border: 1px solid {border_input}; border-radius: 6px;
                padding: 0 10px; background-color: {bg_input};
                color: {text_input}; font-size: 13px;
            }}
            QComboBox:hover {{ border-color: #3b82f6; }}
            QComboBox::drop-down {{ border: none; width: 20px; padding-right: 5px; }}
            QComboBox QAbstractItemView {{
                background-color: {bg_input}; color: {text_input};
                border: 1px solid {border_input};
                selection-background-color: #3b82f6; selection-color: white;
            }}
        """)

        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.sheet_combo, stretch=1)
        self.sheet_selection_widget.hide()
        file_layout.addWidget(self.sheet_selection_widget)

        layout.addWidget(file_section)

        hint_label = QLabel("💡 提示：包含“评估报告”的文件将自动识别为旧报告，另一个为新拆分表。")
        hint_label.setProperty("class", "task-meta")
        hint_color = "#9ca3af" if self._is_dark else "#6b7280"
        hint_label.setStyleSheet(f"color: {hint_color};")
        layout.addWidget(hint_label)
        layout.addStretch()

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        footer = QFrame()
        footer.setObjectName("DialogFooter")
        footer.setFixedHeight(80)
        footer_bg = "#1f2937" if self._is_dark else "transparent"
        footer.setStyleSheet(f"""
            QFrame#DialogFooter {{
                background-color: {footer_bg};
                border-top: 1px solid {border_scroll};
            }}
        """)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(30, 0, 30, 0)

        self.btn_run = QPushButton("开始处理")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setFixedWidth(240)
        self.btn_run.setObjectName("UploadBtn")
        self.btn_run.setStyleSheet("""
            QPushButton#UploadBtn {
                background-color: #3b82f6; color: white; border: none;
                border-radius: 8px; font-size: 14px; font-weight: 600;
            }
            QPushButton#UploadBtn:hover { background-color: #2563eb; }
            QPushButton#UploadBtn:pressed { background-color: #1d4ed8; }
        """)
        self.btn_run.clicked.connect(self.run_process)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_run)
        footer_layout.addStretch()
        main_layout.addWidget(footer)

    def handle_files(self, files):
        expanded_files = []
        for f in files:
            if ArchiveUtils.is_archive(f):
                expanded_files.extend(ArchiveUtils.extract_archive(f))
            else:
                expanded_files.append(f)

        excel_files = [f for f in expanded_files if f.lower().endswith(".xlsx")]
        for f in excel_files:
            filename = os.path.basename(f)
            is_eval = "评估报告" in filename

            if is_eval and not self.excel1_path:
                self.add_file_to_ui(f, is_eval=True)
            elif not is_eval and not self.excel2_path:
                self.add_file_to_ui(f, is_eval=False)
            else:
                if not self.excel1_path:
                    self.add_file_to_ui(f, is_eval=True)
                elif not self.excel2_path:
                    self.add_file_to_ui(f, is_eval=False)

    def add_file_to_ui(self, file_path, is_eval):
        if is_eval:
            if self.excel1_path: return
            self.excel1_path = file_path
        else:
            if self.excel2_path: return
            self.excel2_path = file_path
            # 【新增】当新拆分表上传后，触发后台解析
            self._start_parse_excel2(file_path)

        item = ReReviewFileItem(file_path, is_eval, is_dark=self._is_dark)
        item.remove_clicked.connect(self.remove_file)
        self.queue_layout.insertWidget(self.queue_layout.count() - 1, item)

    def remove_file(self, file_path):
        if self.excel1_path == file_path:
            self.excel1_path = ""
        elif self.excel2_path == file_path:
            self.excel2_path = ""
            self.sheet_selection_widget.hide()
            self.sheet_combo.clear()
            self.excel2_sheet_name = None

        for i in range(self.queue_layout.count()):
            w = self.queue_layout.itemAt(i).widget()
            if isinstance(w, ReReviewFileItem) and w.file_path == file_path:
                w.deleteLater()
                break

    # ================= 【新增】后台解析逻辑 =================
    def _start_parse_excel2(self, file_path):
        """启动后台线程解析新拆分表的工作表"""
        self._parse_worker = ExcelParseWorker("excel2", file_path)
        self._parse_worker.parse_finished.connect(self._on_excel2_parsed)
        self._parse_worker.start()

    def _on_excel2_parsed(self, key, data):
        """新拆分表解析完成回调，填充下拉框并智能选中"""
        if key != "excel2": return
        sheet_names = data.get("sheet_names", [])

        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()

        default_idx = 0
        for i, name in enumerate(sheet_names):
            self.sheet_combo.addItem(f"{i + 1}、{name}", name)
            if "2、" in name and any(kw in name for kw in ["拆分", "功能点"]):
                default_idx = i
            elif any(kw in name for kw in ["功能点拆分表", "拆分表", "功能点"]):
                if default_idx == 0 and i > 0:
                    default_idx = i

        if self.sheet_combo.count() > 0:
            self.sheet_combo.setCurrentIndex(default_idx)
            self.excel2_sheet_name = self.sheet_combo.currentData()
            self.sheet_selection_widget.show()
        self.sheet_combo.blockSignals(False)

    def on_sheet_changed(self, index):
        if index >= 0:
            self.excel2_sheet_name = self.sheet_combo.itemData(index)

    def run_process(self):
        if not self.excel1_path or not self.excel2_path:
            QMessageBox.warning(self, "错误", "请选择两个 Excel 文件（一个旧评估报告，一个新拆分表）")
            return

        import time
        from utils.path_utils import clean_project_name

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        config = MatcherConfig.load()
        base_output_dir = config.get("storage", {}).get("re_review")
        if not base_output_dir:
            base_output_dir = os.path.dirname(self.excel2_path)

        # ============ 统一提取项目名称（只写一次） ============
        e1_full_name = os.path.basename(self.excel1_path)
        e1_full_name = re.sub(r"\.xlsx?$", "", e1_full_name, flags=re.IGNORECASE)
        project_name = re.sub(r"\d{6,8}$", "", e1_full_name)
        project_name = project_name.replace("评估报告", "").strip("-_ ").strip()
        if not project_name:
            project_name = e1_full_name

        # ============ 构建 项目名/时间戳 目录结构 ============
        safe_project_name = clean_project_name(project_name, strip_version=True)
        output_dir = os.path.join(base_output_dir, safe_project_name, timestamp)
        os.makedirs(output_dir, exist_ok=True)
        # ← output_dir 在此处已经确定赋值，后续任何位置引用都不会报错

        # ============ 保存 inputs.json ============
        try:
            inputs_data = {
                "excel1": self.excel1_path,
                "excel2": self.excel2_path,
                "project_name": project_name,
                "timestamp": timestamp,
                "selected_sheet": self.excel2_sheet_name,
            }
            with open(os.path.join(output_dir, "inputs.json"), "w", encoding="utf-8") as f:
                json.dump(inputs_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log_error(f'保存 inputs.json 失败: {e}')

        # ============ 发射任务信号 ============
        try:
            task_info = {
                "project_name": project_name,
                "time": time.strftime("%H:%M:%S"),
                "excel1": self.excel1_path,
                "excel2": self.excel2_path,
                "output_dir": output_dir,
            }
            self.task_started.emit(task_info)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理任务信息时出错: {str(e)}")