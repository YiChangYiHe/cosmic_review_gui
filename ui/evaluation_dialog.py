# ui/evaluation_dialog.py
"""
COSMIC 评估对话框
功能:
1. 上传功能架构图Word文档和评估报告Excel
2. 从功能架构图中读取优化模块
3. 在评估报告的"功能点拆分表"中匹配一二三级模块
4. 在匹配行的备注栏添加"O"
5. 在"结果计算"工作簿的A22填写送审人天
"""

import os
from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QFrame,
    QProgressBar,
    QTextEdit,
    QMessageBox,
    QGroupBox,
    QLineEdit,
    QSpinBox,
    QWidget,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QApplication, QComboBox,
)
from PySide6.QtCore import Qt, Signal, QTimer
from utils.evaluation_processor import EvaluationProcessor
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QWheelEvent
from utils.path_utils import get_resource_path
from utils.runtime_logger import log_error, log_info


class NoWheelComboBox(QComboBox):
    """禁用鼠标滚动的下拉框，防止用户在滚动表格时误触修改动作"""
    def wheelEvent(self, e: QWheelEvent):
        e.ignore()


class ModuleVerificationDialog(QDialog):
    """优化模块匹配确认对话框"""

    def __init__(self, architecture_modules, evaluation_excel_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("确认优化模块匹配情况")
        self.setMinimumSize(700, 550)
        self.architecture_modules = architecture_modules
        self.evaluation_excel_path = evaluation_excel_path

        self._setup_ui()
        # 延迟执行扫描，等窗口显示出来
        QTimer.singleShot(100, self._run_matching_scan)

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        header_layout = QHBoxLayout()
        info_label = QLabel("正在预扫描 Excel 中的优化模块匹配情况：")
        info_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #1e293b;")
        header_layout.addWidget(info_label)

        header_layout.addStretch()
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setVisible(False)
        header_layout.addWidget(self.progress_bar)

        layout.addLayout(header_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["模块编号", "标注名称", "Excel中对应名称", "预览状态"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setStyleSheet("QTableWidget { background-color: white; color: black; }")
        layout.addWidget(self.table)

        tip_label = QLabel("💡 提示：'未找到'通常是因为 Excel 中的名称与架构图不完全一致。您可以确认是否需要继续。")
        tip_label.setStyleSheet("color: #64748b; font-size: 11px;")
        tip_label.setWordWrap(True)
        layout.addWidget(tip_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.cancel_btn = QPushButton("返回修改名称")
        self.cancel_btn.setFixedSize(120, 36)
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.cancel_btn)

        self.confirm_btn = QPushButton("确认无误，开始执行")
        self.confirm_btn.setFixedSize(160, 36)
        self.confirm_btn.setStyleSheet("background-color: #2563eb; color: white; border-radius: 4px; font-weight: bold;")
        self.confirm_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.confirm_btn)

        layout.addLayout(btn_layout)

    def _run_matching_scan(self):
        """执行快速匹配扫描 - 寻找最优匹配结果并记录日志"""
        try:
            import openpyxl
            import re

            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.confirm_btn.setEnabled(False)

            def clean_text(text):
                if not text: return ""
                text = re.sub(r"[（(].*?[)）]", "", str(text))
                return re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]", "", text).lower()

            def calculate_score(s1, s2):
                """计算匹配分值 (0-100)"""
                if not s1 or not s2: return 0
                if s1 == s2: return 100
                if s1 in s2 or s2 in s1: return 90

                set1, set2 = set(s1), set(s2)
                intersection = set1.intersection(set2)
                if not intersection: return 0

                # Jaccard 相似度变形
                score = (len(intersection) * 2.0 / (len(set1) + len(set2))) * 100
                return score

            # 加载 Excel - 耗时操作
            wb = openpyxl.load_workbook(self.evaluation_excel_path, data_only=True, read_only=True)
            ws = EvaluationProcessor.find_sheet_by_names(wb, ["功能点拆分表", "2、功能点拆分表", "拆分表"])

            if not ws:
                QApplication.processEvents()
                log_info('Error: Could not find split sheet in workbook.')
                wb.close()
                self.progress_bar.setVisible(False)
                return

            # 数据池：存储所有非空单元格文本
            excel_pool = []
            for row in ws.iter_rows(min_row=1, min_col=2, max_col=4, values_only=True):
                for cell_val in row:
                    if cell_val:
                        raw = str(cell_val).strip()
                        excel_pool.append({"clean": clean_text(raw), "raw": raw})

            total_tasks = len(self.architecture_modules)
            self.table.setRowCount(total_tasks)
            matched_count = 0

            log_info(f'--- 预扫描启动: 模块总数 {total_tasks} ---')

            for i, (num, title, level, _) in enumerate(self.architecture_modules):
                raw_title = str(title).strip()
                target = clean_text(raw_title)

                best_match = "-"
                best_score = 0

                if target:
                    for item in excel_pool:
                        score = calculate_score(target, item["clean"])
                        if score > best_score:
                            best_score = score
                            best_match = item["raw"]

                # 判定阈值：70分以上视为匹配
                is_matched = best_score >= 70

                self.table.setItem(i, 0, QTableWidgetItem(str(num)))
                self.table.setItem(i, 1, QTableWidgetItem(raw_title))
                self.table.setItem(i, 2, QTableWidgetItem(best_match if is_matched else "-"))

                status_item = QTableWidgetItem()
                if is_matched:
                    matched_count += 1
                    status_item.setText(f"✅ 匹配成功 ({int(best_score)}%)")
                    status_item.setForeground(Qt.darkGreen)
                else:
                    status_item.setText("❌ 未找到")
                    status_item.setForeground(Qt.red)
                    # print(f"未能匹配模块: {raw_title} (最优尝试: {best_match}, 分数: {int(best_score)}%)")

                self.table.setItem(i, 3, status_item)

                # 更新进度条
                progress = int((i + 1) / total_tasks * 100)
                self.progress_bar.setValue(progress)
                if i % 5 == 0: QApplication.processEvents() # 保持界面响应

            log_info(f'--- 预扫描结束: 成功匹配 {matched_count}/{total_tasks} ---')
            wb.close()
            self.confirm_btn.setEnabled(True)
            self.progress_bar.setVisible(False)
        except Exception as e:
            import traceback
            traceback.print_exc()
            log_error(f'Matching preview error: {e}')
            self.confirm_btn.setEnabled(True)
            self.progress_bar.setVisible(False)


class ResultVerificationDialog(QDialog):
    """自动化评估结果确认与修改对话框"""

    def __init__(self, analysis_results, parent=None, source_excel=None):
        super().__init__(parent)
        self.setWindowTitle("确认与修正 COSMIC 评估结果")
        self.setMinimumSize(950, 700)
        self.results = analysis_results
        self.source_excel = source_excel
        self.final_results = []
        self.last_preview_path = None # 记录最近导出的预览文件路径 (用于快捷同步)
        self.redo_requested = False   # 记录是否点击了重新评估

        # 动态获取主题配色
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        self.is_dark = config.get("theme", {}).get("is_dark", False)

        if self.is_dark:
            bg_color = "#1f2937"
            table_bg = "#111827"
            header_bg = "#374151"
            text_color = "#f3f4f6"
            sec_text = "#9ca3af"
            grid_color = "#374151"
            self.item_text_color = "#e5e7eb"
        else:
            bg_color = "#f8fafc"
            table_bg = "white"
            header_bg = "#f1f5f9"
            text_color = "#1e293b"
            sec_text = "#64748b"
            grid_color = "#e2e8f0"
            self.item_text_color = "#1e293b"

        item_text = self.item_text_color
        self.setStyleSheet(f"""
            QDialog {{ background-color: {bg_color}; color: {text_color}; }}
            QTableWidget {{
                background-color: {table_bg};
                border-radius: 8px;
                gridline-color: {grid_color};
                color: {item_text};
                font-size: 13px;
                alternate-background-color: {"#1a222f" if self.is_dark else "#f1f5f9"};
            }}
            QTableWidget::item {{
                color: {item_text};
            }}
            QHeaderView::section {{
                background-color: {header_bg};
                padding: 6px;
                border: 1px solid {grid_color};
                font-weight: bold;
                color: {text_color};
            }}
            QLabel {{
                color: {text_color};
            }}
            QLabel#TitleLabel {{
                font-size: 20px;
                font-weight: bold;
                color: {text_color};
            }}
            QLabel#TipLabel {{
                color: {sec_text};
                font-size: 13px;
            }}
            QComboBox {{
                background-color: {table_bg};
                color: {item_text};
                border: 1px solid {grid_color};
                border-radius: 4px;
                padding: 2px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {table_bg};
                color: {item_text};
                selection-background-color: #2563eb;
            }}
            QPushButton#CancelBtn {{
                background-color: {header_bg};
                color: {text_color};
                border: 1px solid {grid_color};
                border-radius: 6px;
                font-weight: bold;
            }}
            QPushButton#CancelBtn:hover {{
                background-color: {grid_color};
            }}
        """)

        self._setup_ui()
        self._load_data()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # 标题栏
        header = QLabel("🔍 自动化分析结果预览")
        header.setObjectName("TitleLabel")
        layout.addWidget(header)

        tip = QLabel("提示：您可以修改“动作(E/R/X/W)”或“判定结果”，系统将依据您的手动修改进行最终写入。")
        tip.setObjectName("TipLabel")
        layout.addWidget(tip)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["行号", "功能过程", "子过程描述", "AI识别结果", "客户原始类型", "判定结果", "智能分析原因"])

        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)

        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)

        layout.addWidget(self.table)

        # 按钮
        btn_layout = QHBoxLayout()

        # 新增：导出预览和导入修改按钮
        self.export_btn = QPushButton("导出预览 Excel")
        self.export_btn.setFixedSize(140, 40)
        self.export_btn.setStyleSheet("background-color: #0d9488; color: white; border-radius: 6px; font-weight: bold;")
        self.export_btn.clicked.connect(self._on_export_preview)
        btn_layout.addWidget(self.export_btn)

        self.import_btn = QPushButton("同步预览修改")
        self.import_btn.setFixedSize(140, 40)
        self.import_btn.setStyleSheet("background-color: #6366f1; color: white; border-radius: 6px; font-weight: bold;")
        self.import_btn.clicked.connect(self._on_import_modified)
        btn_layout.addWidget(self.import_btn)

        btn_layout.addStretch()

        self.cancel_btn = QPushButton("重新评估")
        self.cancel_btn.setObjectName("CancelBtn")
        self.cancel_btn.setFixedSize(200, 40)
        self.cancel_btn.clicked.connect(self._on_redo_evaluation)
        btn_layout.addWidget(self.cancel_btn)

        self.confirm_btn = QPushButton("确认并写入 Excel")
        self.confirm_btn.setFixedSize(180, 40)
        self.confirm_btn.setStyleSheet("background-color: #2563eb; color: white; border-radius: 6px; font-weight: bold;")
        self.confirm_btn.clicked.connect(self._on_confirm)
        btn_layout.addWidget(self.confirm_btn)

        layout.addLayout(btn_layout)

    def _load_data(self):
        self.table.setRowCount(len(self.results))
        moves = ["E", "R", "X", "W"]  # AI识别结果只能是E/R/X/W
        decisions = ["新增", "复用", "利旧", "优化"]
        from PySide6.QtGui import QColor

        for i, res in enumerate(self.results):
            # 0. 行号
            self.table.setItem(i, 0, QTableWidgetItem(str(res["row"])))
            self.table.item(i, 0).setFlags(Qt.ItemIsEnabled)

            # 1. 功能过程
            proc_text = res.get("group_name", "-")
            self.table.setItem(i, 1, QTableWidgetItem(proc_text))
            self.table.item(i, 1).setFlags(Qt.ItemIsEnabled)
            self.table.item(i, 1).setToolTip(proc_text)

            # 2. 子过程描述
            desc_text = res.get("desc", "无描述")
            desc_item = QTableWidgetItem(desc_text)
            desc_item.setFlags(Qt.ItemIsEnabled)
            desc_item.setToolTip(desc_text)
            self.table.setItem(i, 2, desc_item)

            # 3. AI 识别类型 (Editable NoWheelComboBox) - 只包含E/R/X/W
            ai_move_combo = NoWheelComboBox()
            ai_move_combo.addItems(moves)
            current_ai_move = str(res.get("move", "E")).upper()
            # 如果当前值是n，默认设为E
            if current_ai_move == "N":
                current_ai_move = "E"

            if current_ai_move in moves:
                ai_move_combo.setCurrentText(current_ai_move)
            else:
                ai_move_combo.setCurrentText("E")

            # 连接信号：当下拉框值改变时，自动触发重新评估
            ai_move_combo.currentTextChanged.connect(lambda text, row=i: self._on_move_changed(row))
            self.table.setCellWidget(i, 3, ai_move_combo)

            # 4. 客户原始类型 (Read Only - 显示 Excel 中的原值作为对比)
            orig_move = str(res.get("orig_move", "")).upper()
            move_item = QTableWidgetItem(orig_move if orig_move else "-")
            move_item.setFlags(Qt.ItemIsEnabled)
            move_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 4, move_item)

            # 高亮处理：强制按照用户要求使用黄色高亮显示不一致项
            is_highlight = res.get("is_highlight", False)
            mismatch = (current_ai_move and orig_move and current_ai_move.strip().upper() != orig_move.strip().upper())

            if mismatch or is_highlight:
                # 设置整行背景颜色 (根据深浅色模式调整)
                highlight_color = QColor(255, 255, 180) if not self.is_dark else QColor(60, 50, 0)
                highlight_qss = f"background-color: rgb(255, 255, 180); color: {self.item_text_color};" if not self.is_dark else \
                                f"background-color: rgb(70, 60, 0); color: {self.item_text_color};"

                for col in range(self.table.columnCount()):
                    it = self.table.item(i, col)
                    if it:
                        it.setBackground(highlight_color)
                    # 注意：cellWidget 的背景需要通过 stylesheet 设置
                    w = self.table.cellWidget(i, col)
                    if w:
                        w.setStyleSheet(highlight_qss)

            # 5. 判定结果 (NoWheelComboBox)
            dec_combo = NoWheelComboBox()
            dec_combo.addItems(decisions)

            # 解析原始决策
            orig_val = res.get("value")
            if orig_val == "n":
                current_dec = "利旧"
            elif orig_val == "O":
                current_dec = "优化"
            elif isinstance(orig_val, int):
                current_dec = "复用"
            elif "复用" in str(res.get("reason", "")):
                current_dec = "复用"
            else:
                current_dec = "新增"

            dec_combo.setCurrentText(current_dec)
            self.table.setCellWidget(i, 5, dec_combo)

            # 6. 原因
            orig_l = str(res.get("orig_l", "")).strip()
            l_mismatch = (orig_l != "" and current_dec != orig_l and not (orig_l == "利旧" and current_dec == "利旧"))

            reason_item = QTableWidgetItem(res.get("reason", ""))
            reason_item.setFlags(Qt.ItemIsEnabled)
            reason_item.setToolTip(reason_item.text())
            if l_mismatch or (orig_move and current_ai_move != orig_move):
                reason_item.setForeground(QColor("#dc2626") if not self.is_dark else QColor("#f87171"))
                if l_mismatch:
                    reason_item.setText(f"【变更】{reason_item.text()} (Excel原值: {orig_l})")

            self.table.setItem(i, 6, reason_item)

    def _on_move_changed(self, row_index):
        """当某一行的移动类型改变时，自动重新评估"""
        # 延迟执行，避免频繁触发
        if not hasattr(self, '_eval_timer'):
            self._eval_timer = QTimer()
            self._eval_timer.setSingleShot(True)
            self._eval_timer.timeout.connect(self._on_redo_evaluation)

        # 重置定时器，500ms后自动执行
        self._eval_timer.stop()
        self._eval_timer.start(500)

    def _on_export_preview(self):
        """将当前结果导出为预览 Excel 并直接打开"""
        if not self.source_excel:
            QMessageBox.warning(self, "警告", "未指定源评估文件，无法生成预览。")
            return

        # 1. 采集当前 UI 修改
        self._sync_ui_to_results()

        # 2. 调用后端生成预览
        from utils.evaluation_processor import EvaluationProcessor
        preview_path = EvaluationProcessor.export_preview_excel(self.source_excel, self.results)

        if preview_path:
            # 记录导出路径以备快捷同步 (解决：自动识别预览分析文件)
            self.last_preview_path = preview_path
            # 强化提示：直接打开预览文件
            try:
                os.startfile(preview_path)
                QMessageBox.information(self, "预览已生成",
                    f"预览 Excel 已成功生成并尝试打开：\n{os.path.basename(preview_path)}\n\n您可以在 Excel 中直接修改「AI识别结果」和「判定结果」列，保存后点击“同步预览修改”即可一键同步回传。")
            except Exception as e:
                QMessageBox.information(self, "预览已生成",
                    f"预览文件已生成：\n{preview_path}\n\n自动打开失败，请手动打开修改。")
        else:
            QMessageBox.critical(self, "失败", "生成预览 Excel 失败，请检查文件是否被占用。")

    def _on_import_modified(self):
        """从预览文件重新加载用户的手动修改 (支持快捷自动识别)"""
        import os
        file_path = None

        # 如果存在最近导出的预览文件且文件还在，则直接自动识别，免除重新上传
        if self.last_preview_path and os.path.exists(self.last_preview_path):
            file_path = self.last_preview_path
        else:
            file_path, _ = QFileDialog.getOpenFileName(self, "选择修改后的预览 Excel", "", "Excel Files (*.xlsx *.xlsm)")

        if not file_path:
            return

        from utils.evaluation_processor import EvaluationProcessor
        new_results = EvaluationProcessor.load_results_from_excel(file_path, self.results)

        if new_results:
            self.results = new_results
            self._load_data()
            QMessageBox.information(self, "同步成功", f"分析结果同步成功！\n数据源：{os.path.basename(file_path)}")

    def _on_move_changed(self, row_index):
        """当某一行的移动类型改变时，自动重新评估"""
        # 延迟执行，避免频繁触发
        if not hasattr(self, '_eval_timer'):
            self._eval_timer = QTimer()
            self._eval_timer.setSingleShot(True)
            self._eval_timer.timeout.connect(self._on_redo_evaluation)

        # 重置定时器，500ms后自动执行
        self._eval_timer.stop()
        self._eval_timer.start(500)

    def _sync_ui_to_results(self):
        """辅助方法：将当前表格修改同步到 self.results"""
        for i in range(self.table.rowCount()):
            # 读取 AI 识别动作列
            move_combo = self.table.cellWidget(i, 3)
            if move_combo:
                self.results[i]["move"] = move_combo.currentText()

            # 读取判定结果列
            dec_combo = self.table.cellWidget(i, 5)
            if dec_combo:
                self.results[i]["row_type"] = dec_combo.currentText()

    def _on_redo_evaluation(self):
        """人工修改 AI 识别类型后，点击重新评估刷新结果"""
        # 1. 采集当前 UI 上的所有修改
        self._sync_ui_to_results()

        # 2. 调用后端进行重新推理
        from utils.evaluation_processor import EvaluationProcessor
        new_results = EvaluationProcessor.re_evaluate_inference(self.results)

        # 3. 更新内存数据并刷新界面
        self.results = new_results
        self._load_data()

    def _on_confirm(self):
        self.final_results = []
        for i in range(self.table.rowCount()):
            row_idx = int(self.table.item(i, 0).text())
            # 读取 AI 识别动作列的 ComboBox 值作为最终决定
            move_combo = self.table.cellWidget(i, 3)
            move = move_combo.currentText()

            dec_combo = self.table.cellWidget(i, 5)
            decision_text = dec_combo.currentText()

            res_item = self.results[i].copy()
            res_item["move"] = move
            res_item["row_type"] = decision_text

            # 转换决策为程序可识别格式 (用于写入 Excel)
            if decision_text == "利旧":
                res_item["value"] = "n"
            elif decision_text == "优化":
                res_item["value"] = "O"
            elif decision_text == "新增":
                res_item["value"] = None
            elif decision_text == "复用":
                # 如果是复用，尝试从原始原因中提取复用行号
                orig_reason = str(res_item.get("reason", ""))
                import re
                match = re.search(r"\((\d+)\)", orig_reason)
                if match:
                    res_item["value"] = int(match.group(1))
                else:
                    res_item["value"] = "复用" # 兜底

            self.final_results.append(res_item)

        self.accept()


class FileUploadArea(QFrame):
    """文件上传区域 - 统一样式"""

    file_selected = Signal(str)  # 文件路径

    def __init__(self, title, file_filter, parent=None):
        super().__init__(parent)
        self.title = title
        self.file_filter = file_filter
        self.file_path = None

        self.setProperty("class", "FileUploadArea")
        self.setAcceptDrops(True)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 动态获取主题
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        text_color = "#f3f4f6" if is_dark else "#1e293b"

        # 标题
        title_label = QLabel(self.title)
        title_label.setStyleSheet(
            f"font-size: 13px; font-weight: bold; margin-bottom: 4px; color: {text_color};"
        )
        layout.addWidget(title_label)

        # 文件路径显示
        self.path_label = QLabel("未选择文件")
        self.path_label.setProperty("class", "task-meta")
        self.path_label.setStyleSheet(f"color: {text_color};")
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

        # 按钮
        btn_layout = QHBoxLayout()
        self.select_btn = QPushButton("选择文件")
        self.select_btn.setProperty("class", "SecondaryBtn")
        self.select_btn.setFixedHeight(36)
        self.select_btn.setStyleSheet(f"color: {text_color};")
        self.select_btn.clicked.connect(self._select_file)
        btn_layout.addWidget(self.select_btn)

        self.clear_btn = QPushButton("清除")
        self.clear_btn.setProperty("class", "SecondaryBtn")
        self.clear_btn.setFixedHeight(36)
        self.clear_btn.setStyleSheet(f"color: {text_color};")
        self.clear_btn.clicked.connect(self._clear_file)
        self.clear_btn.setEnabled(False)
        btn_layout.addWidget(self.clear_btn)

        # 标注按钮 (仅图片模式可用)
        self.annotate_btn = QPushButton("🖱️ 进入标注")
        self.annotate_btn.setProperty("class", "PrimaryBtn")
        self.annotate_btn.setFixedHeight(36)
        self.annotate_btn.setVisible(False)
        btn_layout.addWidget(self.annotate_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 统一样式
        self.setStyleSheet(
            f"""
            QFrame[class="FileUploadArea"] {{
                border: 2px dashed #4b5563;
                border-radius: 8px;
                background-color: {"rgba(31, 41, 55, 0.3)" if is_dark else "#f8fafc"};
            }}
            QFrame[class="FileUploadArea"]:hover {{
                border-color: #3b82f6;
                background-color: {"rgba(59, 130, 246, 0.1)" if is_dark else "#f1f5f9"};
            }}
        """
        )

    def _select_file(self):
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        initial_dir = config.get("evaluation_folder", "")
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = ""

        file_path, _ = QFileDialog.getOpenFileName(
            self, f"选择{self.title}", initial_dir, self.file_filter
        )
        if file_path:
            self.set_file(file_path)

    def _clear_file(self):
        self.file_path = None
        self.path_label.setText("未选择文件")
        self.clear_btn.setEnabled(False)
        self.annotate_btn.setVisible(False)
        self.file_selected.emit("")

    def clear_file(self):
        """外部清除方法"""
        self._clear_file()

    def set_file(self, file_path):
        """设置文件路径"""
        if os.path.exists(file_path):
            self.file_path = file_path
            self.path_label.setText(f"✓ {os.path.basename(file_path)}")
            self.clear_btn.setEnabled(True)
            # 标记为图片时显示标注按钮
            if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                self.annotate_btn.setVisible(True)
            # 发射信号，通知父组件路径已更新
            self.file_selected.emit(file_path)
            # 如果是图片文件，显示标注按钮
            if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                self.annotate_btn.setVisible(True)
            self.file_selected.emit(file_path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            # 验证文件类型
            if self.file_filter == "图片文件 (*.png *.jpg *.jpeg *.bmp)":
                if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp")):
                    self.set_file(file_path)
            elif self.file_filter == "Excel文件 (*.xlsx *.xls)":
                if file_path.lower().endswith((".xlsx", ".xls")):
                    self.set_file(file_path)


class EvaluationDialog(QDialog):
    """COSMIC评估对话框"""

    # 定义开始评估信号
    evaluation_requested = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COSMIC 功能点评估")
        self.setMinimumSize(900, 750)
        self.resize(950, 800)

        # 设置透明背景
        self.setAttribute(Qt.WA_TranslucentBackground, False)

        # 文件路径
        self.architecture_doc_path = None
        self.evaluation_excel_path = None
        self.annotated_modules = []  # 存储标注的模块列表

        self._setup_ui()

    def _setup_ui(self):
        # 应用原生标题栏深色模式
        from PySide6.QtWidgets import QApplication
        from utils.styles import apply_dark_title_bar

        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )
        apply_dark_title_bar(self, is_dark)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 滚动区域
        from PySide6.QtWidgets import QScrollArea
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet(
            """
            QScrollArea {
                border: none;
                background: transparent;
            }
        """
        )

        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(30, 30, 30, 20)
        layout.setSpacing(20)

        # 标题
        title = QLabel("📊 COSMIC 功能点评估")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet(f"color: {'#f3f4f6' if is_dark else '#1e293b'};")
        layout.addWidget(title)

        # 说明文本
        desc = QLabel(
            "💡 本工具将:\n"
            "   • 从功能架构图中提取优化模块 (交互标注或手动输入)\n"
            "   • 在评估报告的'功能点拆分表'中匹配一二三级模块并标注'O'\n"
            "   • 在'结果计算'工作簿的A22填写送审人天\n"
            "   • 支持COSMIC规则评估"
        )
        desc.setProperty("class", "task-meta")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 输入方式选择
        mode_section = QFrame()
        mode_section.setProperty("class", "ContentSection")
        mode_layout = QVBoxLayout(mode_section)
        mode_layout.setContentsMargins(20, 20, 20, 20)
        mode_layout.setSpacing(12)

        mode_title = QLabel("📝 输入方式")
        mode_title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {'#f3f4f6' if is_dark else '#1e293b'};")
        mode_layout.addWidget(mode_title)

        # 模式选择按钮
        btn_layout = QHBoxLayout()
        self.manual_radio = QPushButton("✍️ 手动输入 (推荐)")
        self.manual_radio.setCheckable(True)
        self.manual_radio.setChecked(True)  # 默认手动输入
        self.manual_radio.setProperty("class", "ModeBtn")
        self.manual_radio.setFixedHeight(40)
        self.manual_radio.clicked.connect(self._on_mode_changed)
        btn_layout.addWidget(self.manual_radio)

        self.annotate_radio = QPushButton("🖱️ 交互标注")
        self.annotate_radio.setCheckable(True)
        self.annotate_radio.setChecked(False)
        self.annotate_radio.setProperty("class", "ModeBtn")
        self.annotate_radio.setFixedHeight(40)
        self.annotate_radio.clicked.connect(self._on_mode_changed)
        btn_layout.addWidget(self.annotate_radio)
        mode_layout.addLayout(btn_layout)

        # 标注说明（仅标注模式显示）
        self.annotate_tip = QLabel("💡 上传功能架构图后，框选优化模块区域，系统自动OCR识别文字并选择层级")
        self.annotate_tip.setProperty("class", "task-meta")
        self.annotate_tip.setStyleSheet("color: #10b981; padding: 8px; background-color: rgba(16, 185, 129, 0.1); border-radius: 4px;")
        self.annotate_tip.setWordWrap(True)
        self.annotate_tip.setVisible(False)  # 默认隐藏
        mode_layout.addWidget(self.annotate_tip)

        layout.addWidget(mode_section)

        # 文件上传区域
        upload_section = QFrame()
        upload_section.setProperty("class", "ContentSection")
        upload_layout = QVBoxLayout(upload_section)
        upload_layout.setContentsMargins(20, 20, 20, 20)
        upload_layout.setSpacing(12)

        upload_title = QLabel("📁 文件上传")
        upload_title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {'#f3f4f6' if is_dark else '#1e293b'};")
        upload_layout.addWidget(upload_title)

        # 功能架构图上传 (标注模式) - 支持图片、Word 和 Excel 拆分表
        self.architecture_upload = FileUploadArea(
            "参考源文件 (Word架构图/Excel拆分表/图片)", "所有支持的文件 (*.png *.jpg *.jpeg *.bmp *.docx *.doc *.xlsx);;图片文件 (*.png *.jpg *.jpeg *.bmp);;Word文档 (*.docx *.doc);;Excel文件 (*.xlsx)"
        )
        self.architecture_upload.file_selected.connect(self._on_architecture_selected)
        self.architecture_upload.annotate_btn.clicked.connect(
            lambda: self._open_annotator(self.architecture_doc_path)
        )
        # 即使在手动模式也允许上传源文件，作为 A-K 搬运的来源
        self.architecture_upload.setVisible(True)
        upload_layout.addWidget(self.architecture_upload)

        # 手动输入区域 (手动模式)
        self.manual_input_widget = QWidget()
        manual_layout = QVBoxLayout(self.manual_input_widget)
        manual_layout.setContentsMargins(0, 0, 0, 0)

        manual_label = QLabel("💡 提示：如果上传了 Excel 拆分表，系统将自动搬运 A-K 列数据；否则请在此输入优化模块列表。")
        manual_label.setProperty("class", "task-meta")
        manual_layout.addWidget(manual_label)

        self.manual_input_text = QTextEdit()
        self.manual_input_text.setPlaceholderText(
            "示例:\n"
            "4.1 用户管理优化\n"
            "4.2 数据查询优化\n"
            "4.2.1 查询接口优化\n"
            "4.2.2 缓存机制优化"
        )
        self.manual_input_text.setMaximumHeight(150)
        self.manual_input_text.textChanged.connect(self._check_ready)
        manual_layout.addWidget(self.manual_input_text)

        self.manual_input_widget.setVisible(True)  # 默认显示
        upload_layout.addWidget(self.manual_input_widget)

        # 评估报告上传
        self.evaluation_upload = FileUploadArea(
            "评估报告 (Excel)", "Excel文件 (*.xlsx *.xls)"
        )
        self.evaluation_upload.file_selected.connect(self._on_evaluation_selected)
        upload_layout.addWidget(self.evaluation_upload)

        layout.addWidget(upload_section)

        # 送审人天设置
        param_section = QFrame()
        param_section.setProperty("class", "ContentSection")
        param_layout = QVBoxLayout(param_section)
        param_layout.setContentsMargins(20, 20, 20, 20)
        param_layout.setSpacing(12)

        param_title = QLabel("⚙️ 评估参数")
        param_title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {'#f3f4f6' if is_dark else '#1e293b'};")
        param_layout.addWidget(param_title)

        manday_layout = QHBoxLayout()
        manday_label = QLabel("送审人天:")
        manday_label.setProperty("class", "task-meta")
        manday_layout.addWidget(manday_label)

        self.manday_input = QSpinBox()
        self.manday_input.setMinimum(0)
        self.manday_input.setMaximum(9999)
        self.manday_input.setValue(0)
        self.manday_input.setSuffix(" 人天")
        self.manday_input.setFixedSize(150, 36)
        manday_layout.addWidget(self.manday_input)

        manday_layout.addStretch()
        param_layout.addLayout(manday_layout)

        # 新增：自动评估开关
        self.auto_eval_check = QPushButton("🤖 开启自动评估规则 (实验性)")
        self.auto_eval_check.setCheckable(True)
        self.auto_eval_check.setChecked(True)
        self.auto_eval_check.setProperty("class", "ModeBtn")
        self.auto_eval_check.setToolTip("根据功能描述自动分析E,X,R,W及变更类型")
        self.auto_eval_check.setFixedHeight(36)
        self.auto_eval_check.setFixedWidth(250)
        param_layout.addWidget(self.auto_eval_check)

        layout.addWidget(param_section)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(6)
        layout.addWidget(self.progress_bar)

        # 日志区域
        log_section = QFrame()
        log_section.setProperty("class", "ContentSection")
        log_layout = QVBoxLayout(log_section)
        log_layout.setContentsMargins(20, 20, 20, 20)
        log_layout.setSpacing(12)

        log_title = QLabel("📋 处理日志")
        log_title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {'#f3f4f6' if is_dark else '#1e293b'};")
        log_layout.addWidget(log_title)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet(
            f"""
            QTextEdit {{
                border: 1px solid #4b5563;
                border-radius: 6px;
                padding: 8px;
                background-color: {"rgba(31, 41, 55, 0.3)" if is_dark else "#f9fafb"};
                color: {"#f3f4f6" if is_dark else "#111827"};
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 11px;
            }}
        """
        )
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_section)

        layout.addStretch()
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        # 底部按钮区域
        footer = QFrame()
        footer.setObjectName("DialogFooter")
        footer.setFixedHeight(80)
        footer.setStyleSheet(
            """
            QFrame#DialogFooter {
                background-color: transparent;
                border-top: 1px solid #374151;
            }
        """
        )
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(30, 0, 30, 0)

        # 关闭按钮
        self.cancel_btn = QPushButton("关闭")
        self.cancel_btn.setProperty("class", "secondary-btn")
        self.cancel_btn.setFixedSize(100, 45)
        self.cancel_btn.clicked.connect(self.reject)
        footer_layout.addWidget(self.cancel_btn)

        footer_layout.addStretch()

        # 开始评估按钮
        self.start_btn = QPushButton("开始评估")
        self.start_btn.setProperty("class", "UploadBtn")
        self.start_btn.setFixedSize(240, 45)
        self.start_btn.clicked.connect(self._start_evaluation)
        self.start_btn.setEnabled(False)
        footer_layout.addWidget(self.start_btn)
        footer_layout.addStretch()

        main_layout.addWidget(footer)

    def _on_mode_changed(self):
        """切换输入模式"""
        # 互斥选择
        if self.sender() == self.manual_radio:
            self.manual_radio.setChecked(True)
            self.annotate_radio.setChecked(False)
            is_manual = True
        else:
            self.annotate_radio.setChecked(True)
            self.manual_radio.setChecked(False)
            is_manual = False

        # 显示/隐藏对应区域
        self.manual_input_widget.setVisible(is_manual)
        # self.architecture_upload 保持可见，因为它可以作为 Excel 搬运的源
        self.annotate_tip.setVisible(not is_manual)

        # 自动触发：如果切换到交互标注，且已经选择了图片，自动弹出标注页
        if not is_manual and self.architecture_doc_path:
            self._open_annotator(self.architecture_doc_path)

        # 重新检查是否可以开始
        self._check_ready()

    def _on_architecture_selected(self, file_path):
        """功能架构图选择"""
        self.architecture_doc_path = file_path if file_path else None

        if file_path:
            filename = os.path.basename(file_path)
            self._log(f"已选择参考源文件: {filename}")

            # 检测文件类型并提供反馈
            if filename.lower().endswith(('.xlsx', '.xls')):
                self._log("💡 识别为 Excel 拆分表，评估时将自动迁移该文件的 A-K 列数据。")

                # 自动寻找评估报告模板 (folder 目录下)
                if not self.evaluation_excel_path:
                    try:
                        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                        folder_dir = os.path.join(base_dir, "folder")
                        if os.path.exists(folder_dir):
                            templates = [f for f in os.listdir(folder_dir) if "评估报告" in f and f.endswith(".xlsx")]
                            if templates:
                                template_path = os.path.join(folder_dir, templates[0])
                                self.evaluation_upload.set_file(template_path)
                                self._log(f"✨ 已为您自动匹配评估报告模板: {templates[0]}")
                    except Exception as e:
                        self._log(f"⚠️ 自动寻找模板失败: {str(e)}")

            elif filename.lower().endswith(('.docx', '.doc')):
                self._log("💡 识别为 Word 架构图，由于暂不支持直接解析 Word 表格，建议手动输入优化模块。")
                self._handle_word_architecture(file_path)
            elif filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                self._log("💡 识别为图片架构图，您可以点击下方“进入标注”进行交互式识别。")
                if self.annotate_radio.isChecked():
                    self._open_annotator(file_path)

        self._check_ready()

    def _handle_word_architecture(self, file_path):
        """处理 Word 中的架构图提取"""
        self._log("🔍 正在从 Word 中提取架构图...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(20)

        try:
            from utils.document_processor import DocumentProcessor
            # 提取图片
            image_paths = DocumentProcessor.extract_images_from_word(file_path)

            if not image_paths:
                self._log("⚠️ 未在 Word 中发现图片。")
                return

            self._log(f"✅ 成功从 Word 提取 {len(image_paths)} 张图片。")

            # 使用列表选择器让用户选择正确的架构图
            from ui.upload_dialog import ImageSelectionDialog
            selector = ImageSelectionDialog(image_paths, self)
            if selector.exec() == QDialog.Accepted:
                selected_path = selector.selected_image
                if selected_path:
                    self.architecture_doc_path = selected_path
                    self._log(f"已选择架构图: {os.path.basename(selected_path)}")
                    # 切换到标注
                    if self.annotate_radio.isChecked():
                        self._open_annotator(selected_path)
        except Exception as e:
            self._log(f"❌ 提取 Word 图片失败: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def _open_annotator(self, image_path):
        """打开图像标注对话框"""
        try:
            from utils.image_annotator import ImageAnnotatorDialog

            # 传入之前保存的标注结果（如果有）
            dialog = ImageAnnotatorDialog(image_path, self, initial_annotations=self.annotated_modules)
            if dialog.exec() == QDialog.Accepted:
                # 获取标注的完整模块数据 [(rect, num, name, color), ...]
                modules = dialog.get_modules()
                if modules:
                    # 保存标注结果供后续使用
                    self.annotated_modules = modules
                    self._log(f"标注完成，共标注 {len(modules)} 个优化模块")
                else:
                    self._log("未标注任何模块")
                    self.annotated_modules = []
            else:
                # 用户取消了标注，不清除已有标注，只是不更新
                self._log("已取消标注编辑")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开标注工具失败:\n{str(e)}")
            self._log(f"错误: {str(e)}")

    def _on_evaluation_selected(self, file_path):
        """评估报告选择"""
        self.evaluation_excel_path = file_path if file_path else None
        self._check_ready()
        if file_path:
            self._log(f"已选择评估报告: {os.path.basename(file_path)}")

    def _check_ready(self):
        """检查是否可以开始评估"""
        is_evaluation_ready = self.evaluation_excel_path is not None

        # 允许不标注架构图直接评估 (针对只需评估新增或无优化模块的项目)
        self.start_btn.setEnabled(is_evaluation_ready)

    def _log(self, message):
        """添加日志"""
        self.log_text.append(message)

    def _update_progress(self, value, message=""):
        """更新进度"""
        self.progress_bar.setValue(value)
        if message:
            self._log(message)

    def _start_evaluation(self):
        """开始评估 - 发送信号给主窗口并立即关闭"""
        try:
            # 获取输入模式
            is_manual = self.manual_radio.isChecked()

            # 准备输入数据
            if is_manual:
                manual_text = self.manual_input_text.toPlainText().strip()
                architecture_input = manual_text
            else:
                # 将标注结果转换为文本格式（后端处理器需要文本格式匹配）
                # 如果没有标注任何模块，architecture_input 将为空字符串，这是允许的
                architecture_input = "\n".join([f"{num} {name}" for rect, num, name, color in self.annotated_modules])

            if not os.path.exists(self.evaluation_excel_path):
                QMessageBox.warning(self, "错误", "评估报告文件不存在")
                return

            # --- 新增：弹出匹配确认对话框 ---
            from utils.evaluation_processor import EvaluationProcessor
            # 解析模块列表
            temp_modules = EvaluationProcessor.parse_manual_input(architecture_input)
            if temp_modules:
                verify_dialog = ModuleVerificationDialog(temp_modules, self.evaluation_excel_path, self)
                if verify_dialog.exec() != QDialog.Accepted:
                    return # 用户选择返回修改
            # ---------------------------

            # 智能提取项目名称
            excel_name = os.path.basename(self.evaluation_excel_path).replace(".xlsx", "").replace(".xls", "")
            import re
            project_name = re.sub(r"^附件\s*\d+\s*", "", excel_name)
            project_name = re.sub(r"功能点拆分表.*$", "", project_name)

            # 如果清理后包含 xxxx 或太短，尝试从文件夹路径获取
            if "xxxx" in project_name.lower() or len(project_name) < 3:
                from extend.matcher_config import MatcherConfig
                cfg = MatcherConfig.load()
                folder = cfg.get("evaluation_folder", "")
                if folder:
                    folder_name = os.path.basename(folder)
                    # 清理文件夹名中的日期或编号 e.g. 【260128】
                    folder_name = re.sub(r"【.*?】", "", folder_name)
                    project_name = folder_name.strip()

            # 构造任务信息
            task_info = {
                "task_id": datetime.now().strftime("%Y%m%d%H%M%S"),
                "project_name": project_name,
                "architecture_input": architecture_input,
                "evaluation_excel_path": self.evaluation_excel_path,
                "manday": self.manday_input.value(),
                "is_auto_evaluate": self.auto_eval_check.isChecked(),
                "architecture_file_path": self.architecture_doc_path,
                "architecture_file": os.path.basename(self.architecture_doc_path) if self.architecture_doc_path else "手动输入",
                "status": "评估中",
                "create_time": datetime.now()
            }

            # 发送信号并关闭
            self.evaluation_requested.emit(task_info)
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"启动评估失败:\n{str(e)}")

    def _on_finished(self, output_path):
        """评估完成"""
        self._log("\n" + "=" * 50)
        self._log(f"✓ 评估完成!")
        self._log(f"输出文件: {output_path}")

        QMessageBox.information(
            self, "完成", f"COSMIC评估已完成!\n\n输出文件:\n{output_path}"
        )
        self._reset_ui()

    def _on_error(self, error_msg):
        """评估出错"""
        self._log(f"\n✗ 错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"评估过程出错:\n{error_msg}")
        self._reset_ui()

    def _reset_ui(self):
        """重置UI状态"""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
