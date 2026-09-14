from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QPushButton,
    QFrame,
    QTextBrowser,
    QApplication,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QGuiApplication, QIcon
from utils.path_utils import clean_project_name, get_resource_path
from utils.styles import apply_dark_title_bar
import os
import sys

# 必须在进程最开始设置，否则 DPI 可能失效
if sys.platform == "win32":
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_AUTOSCREENSCALEFACTOR"] = "1"


class SummaryDialog(QDialog):
    """结果情况汇总弹窗 - 汇总后5个步骤的结果"""

    def __init__(self, task_data=None, parent=None, summary_html=None):
        # 确保从外部 (如 tkinter) 调用时，QApplication 已初始化且开启 DPI
        if not QApplication.instance():
            QApplication.setHighDpiScaleFactorRoundingPolicy(
                Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
            )
            self._internal_app = QApplication(sys.argv or [])
            self._internal_app.setFont(QFont("Microsoft YaHei UI", 9))

        super().__init__(parent)
        self.task_data = task_data or {}

        # 应用原生标题栏深色模式
        is_dark = "background-color: #1f2937" in (
                QApplication.instance().styleSheet() or ""
        )
        apply_dark_title_bar(self, is_dark)

        self.setWindowTitle("审核结果情况汇总")
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.ico")))
        self.setMinimumSize(520, 480)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title_label = QLabel(" 审核结果汇总")
        title_label.setStyleSheet(
            "font-size: 18px; font-weight: bold; font-family: 'Microsoft YaHei UI';"
        )
        layout.addWidget(title_label)

        # 改用 QTextBrowser
        self.text_browser = QTextBrowser()
        self.text_browser.setReadOnly(True)

        # 【核心修复】：决定使用哪个 HTML
        if summary_html:
            # 优先使用外部传入的已保存 HTML（例如历史页面直接读取的）
            final_html = summary_html
        else:
            # 否则调用模块级别的独立函数动态生成
            final_html = generate_summary_html(self.task_data)

        self.text_browser.setHtml(final_html)
        layout.addWidget(self.text_browser)

        # 按钮区域
        btn_layout = QHBoxLayout()

        copy_btn = QPushButton("📋 复制全部")
        copy_btn.setFixedHeight(35)
        copy_btn.setStyleSheet(
            """
            QPushButton { background: #2563eb; color: white; border-radius: 6px; font-weight: bold; padding: 0 15px; font-family: 'Microsoft YaHei UI'; }
            QPushButton:hover { background: #1d4ed8; }
        """
        )
        copy_btn.clicked.connect(self._copy_to_clipboard)

        close_btn = QPushButton("关闭")
        close_btn.setFixedHeight(35)
        close_btn.setStyleSheet(
            """
            QPushButton { border: 1px solid palette(mid); border-radius: 6px; padding: 0 15px; font-family: 'Microsoft YaHei UI'; }
            QPushButton:hover { background: palette(alternate-base); }
        """
        )
        close_btn.clicked.connect(self.accept)

        btn_layout.addStretch()

        # 添加打开辅助报告的入口 (辅助报告路径依然从 task_data 中获取)
        results_list = self.task_data.get("validation_results", [])
        if results_list:
            v_res = results_list[0]

            # 层级匹配报告
            h_path = v_res.get("hierarchy_res", {}).get("report_path")
            if h_path and os.path.exists(h_path):
                h_btn = QPushButton(" 层级匹配报告")
                h_btn.setFixedHeight(35)
                h_btn.setStyleSheet(
                    "border: 1px solid #3b82f6; color: #3b82f6; border-radius: 6px; padding: 0 10px;"
                )
                h_btn.clicked.connect(lambda: os.startfile(h_path))
                btn_layout.addWidget(h_btn)

            # 功能过程报告
            p_path = v_res.get("process_res", {}).get("report_path")
            if p_path and os.path.exists(p_path):
                p_btn = QPushButton("📁 功能过程报告")
                p_btn.setFixedHeight(35)
                p_btn.setStyleSheet(
                    "border: 1px solid #f59e0b; color: #f59e0b; border-radius: 6px; padding: 0 10px;"
                )
                p_btn.clicked.connect(lambda: os.startfile(p_path))
                btn_layout.addWidget(p_btn)

            # 数据移动类型报告
            m_path = v_res.get("move_res", {}).get("report_path")
            if m_path and os.path.exists(m_path):
                m_btn = QPushButton("📁 数据移动报告")
                m_btn.setFixedHeight(35)
                m_btn.setStyleSheet(
                    "border: 1px solid #10b981; color: #10b981; border-radius: 6px; padding: 0 10px;"
                )
                m_btn.clicked.connect(lambda: os.startfile(m_path))
                btn_layout.addWidget(m_btn)

        btn_layout.addWidget(copy_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def _copy_to_clipboard(self):
        clipboard = QGuiApplication.clipboard()
        clipboard.setText(self.text_browser.toPlainText())
        # 变色反馈
        source_btn = self.sender()
        if source_btn:
            source_btn.setText("✅ 已复制")
            source_btn.setStyleSheet(
                "background: #10b981; color: white; border-radius: 6px; font-weight: bold; padding: 0 15px; font-family: 'Microsoft YaHei UI';"
            )


class ReportDialog(QDialog):
    """详细报告弹窗 - 模板校验专用"""

    def __init__(self, task_data, parent=None):
        super().__init__(parent)
        self.task_data = task_data

        # 应用原生标题栏深色模式
        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )
        apply_dark_title_bar(self, is_dark)

        self.setWindowTitle("Cosmic 智能审核报告 - 模板校验详情")
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.ico")))
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        # 净化项目名称
        clean_name = clean_project_name(task_data["filename"])

        header = QLabel(f"📋 项目名称: {clean_name}")
        header.setStyleSheet(
            "font-size: 20px; font-weight: bold; font-family: 'Microsoft YaHei UI';"
        )
        layout.addWidget(header)

        # 统计卡片
        stats_layout = QHBoxLayout()
        # 修正：从 raw_task_info 中获取校验结果
        # 修正：直接从 task_data 获取校验结果 (由 TaskCard 填充)
        v_res_list = task_data.get("validation_results", [])
        v_res = v_res_list[0] if v_res_list else {}

        stats_layout.addWidget(
            self._create_stat_card(
                "总体匹配率",
                f"{v_res.get('template_match_rate', 0)*100:.1f}%",
                "#2563eb",
            )
        )
        stats_layout.addWidget(
            self._create_stat_card(
                "已匹配", str(len(v_res.get("matched_modules", []))), "#10b981"
            )
        )
        stats_layout.addWidget(
            self._create_stat_card(
                "疑似项", str(len(v_res.get("suspect_modules", []))), "#ef4444"
            )
        )
        stats_layout.addWidget(
            self._create_stat_card(
                "缺失项", str(len(v_res.get("missing_modules", []))), "#ef4444"
            )
        )
        layout.addLayout(stats_layout)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            [
                "📊 层级",
                "📚 标准章节",
                "⚖️ 判定结果",
                "📝 模板参考句 (对照)",
                "📤 上传文本句 (对照)",
                "💯 重合度",
            ]
        )
        # 移除硬编码样式
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(5, 80)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setWordWrap(True)

        self._fill_table(v_res)
        self.table.resizeRowsToContents()
        layout.addWidget(self.table)

        # 底部关闭按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        # 新增项目：打开最终报告按钮 (按图一蓝色框位置)
        self.open_excel_btn = QPushButton("📂 打开最终报告")
        self.open_excel_btn.setFixedSize(140, 40)
        self.open_excel_btn.setStyleSheet(
            """
            QPushButton { border: 1px solid #10b981; color: #10b981; border-radius: 6px; font-weight: bold; }
            QPushButton:hover { background: #10b981; color: white; }
        """
        )
        self.open_excel_btn.clicked.connect(self._open_excel_report)
        btn_layout.addWidget(self.open_excel_btn)

        close_btn = QPushButton("确定")
        close_btn.setFixedSize(100, 40)
        close_btn.setStyleSheet(
            """
            QPushButton { background: #2563eb; color: white; border-radius: 6px; font-weight: bold; }
            QPushButton:hover { background: #1d4ed8; }
        """
        )
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def _open_excel_report(self):
        """打开生成的 Excel 报表"""
        import os

        v_res_list = self.task_data.get("validation_results", [])
        if not v_res_list:
            return
        path = v_res_list[0].get("auto_report_path")
        if path and os.path.exists(path):
            os.startfile(os.path.abspath(path))

    def _create_stat_card(self, label, value, color):
        card = QFrame()
        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )

        # 映射标签到图标
        icon_map = {"总体匹配率": "📈", "已匹配": "✅", "疑似项": "🔍", "缺失项": "❌"}
        icon = icon_map.get(label, "📊")

        if is_dark:
            # 深色模式：更纯净的黑底白字
            card.setStyleSheet(
                f"QFrame {{ background: #000000; border-radius: 12px; border: 1px solid #374151; padding: 15px; }}"
            )
            label_style = "color: #ffffff; font-size: 13px; font-family: 'Microsoft YaHei UI'; font-weight: bold;"
            value_style = f"color: #ffffff; font-size: 26px; font-weight: bold; font-family: 'Microsoft YaHei UI';"  # 强制白色数值
        else:
            # 浅色模式
            card.setStyleSheet(
                f"QFrame {{ background: #f8fafc; border-radius: 12px; border: 1px solid #e2e8f0; padding: 15px; }}"
            )
            label_style = (
                "color: #64748b; font-size: 13px; font-family: 'Microsoft YaHei UI';"
            )
            value_style = f"color: {color}; font-size: 26px; font-weight: bold; font-family: 'Microsoft YaHei UI';"

        layout = QVBoxLayout(card)
        l1 = QLabel(f"{icon} {label}")
        l1.setStyleSheet(label_style)

        v1 = QLabel(value)
        # 如果不是深色模式下的强制白色，则使用传入的颜色
        if is_dark:
            v1.setStyleSheet(value_style)
        else:
            v1.setStyleSheet(value_style)

        layout.addWidget(l1)
        layout.addWidget(v1)
        return card
        layout.addWidget(l1)
        layout.addWidget(v1)
        return card

    def _fill_table(self, v_res):
        details = v_res.get("all_details", [])
        self.table.setRowCount(len(details))

        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )
        default_color = "#f9fafb" if is_dark else "#1e293b"

        tag_styles = {
            "🚨 雷同": ("🚨 雷同", "#ef4444"),
            "🔹 参考": ("🔹 参考", "#2563eb"),
            "✅ 原创": ("✅ 原创", "#10b981"),
            "🔹 模板参考": ("🔹 模板余留", "#64748b"),
            "缺失": ("❌ 缺失", "#ef4444"),
            "已匹配": ("✅ 已匹配", "#10b981"),
        }

        for row, d in enumerate(details):
            level = d.get("level", 1)

            # 第一列：层级
            lvl_item = QTableWidgetItem(str(level))
            lvl_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, lvl_item)

            # 第二列：章节 (带缩进)
            indent = "    " * (level - 1)
            chap_item = QTableWidgetItem(f"{indent}{d.get('chapter', '-')}")
            if level == 1:
                f = chap_item.font()
                f.setBold(True)
                chap_item.setFont(f)
            self.table.setItem(row, 1, chap_item)

            # 第三列：判定结果 (美化图标)
            status = d.get("status", "")
            display_text, color = tag_styles.get(status, (status, default_color))
            st_item = QTableWidgetItem(display_text)
            st_item.setForeground(QColor(color))
            st_item.setTextAlignment(Qt.AlignCenter)
            st_font = QFont()
            st_font.setBold(True)
            st_item.setFont(st_font)
            self.table.setItem(row, 2, st_item)

            # 第四列：模板参考句
            tpl_s = d.get("tpl_sentence", "-")
            tpl_item = QTableWidgetItem(tpl_s)
            tpl_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(row, 3, tpl_item)

            # 第五列：上传文本句
            tgt_s = d.get("target_sentence", "-")
            tgt_item = QTableWidgetItem(tgt_s)
            tgt_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(row, 4, tgt_item)

            # 第六列：重合度
            score_val = d.get("score", 0)
            score_text = f"{score_val*100:.1f}%" if status != "缺失" else "-"
            score_item = QTableWidgetItem(score_text)
            score_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 5, score_item)

def generate_summary_html(task_data):
    """独立生成汇总 HTML，供后台保存和历史页面直接读取"""
    from PySide6.QtWidgets import QApplication
    is_dark = "background-color: #1f2937" in (QApplication.instance().styleSheet() or "")
    title_color = "#f9fafb" if is_dark else "#1e293b"
    text_color = "#d1d5db" if is_dark else "#475569"
    border_color = "#374151" if is_dark else "#f1f5f9"

    filename = clean_project_name(task_data.get("filename", "未知项目"))
    results_list = task_data.get("validation_results", [])

    html = f"""
    <div style="font-family: 'Microsoft YaHei UI'; color: {text_color}; line-height: 1.8;">
        <div style="font-size: 20px; font-weight: bold; color: {title_color}; margin-bottom: 25px; white-space: nowrap; border-bottom: 1px solid {border_color}; padding-bottom: 10px;">
            {filename}
        </div>
    """
    if not results_list:
        html += "（尚未获取到校验结果）</div>"
        return html

    results = results_list[0]
    points = []

    # 1. 送审比例
    ratio_res = results.get("ratio_check", {})
    if ratio_res and not ratio_res.get("is_ok"):
        ratio = ratio_res.get("ratio", 0)
        mandays = ratio_res.get("mandays", 0)
        target = 2.0 if mandays <= 1000 else 1.5
        trend = "过多" if ratio >= target else "过少"
        points.append(f"送审比例{trend}，送审比例在 0.8 &lt; 送审功能点/送审人天 &lt; {target}")

    # 2. Excel 空值检查
    excel_check = results.get("excel_check", {})
    if excel_check and not excel_check.get("is_ok"):
        for err in excel_check.get("errors", [])[:10]:
            points.append(err)

    # 3. 层级匹配
    hierarchy_res = results.get("hierarchy_res", {})
    combined_h = hierarchy_res.get("not_found_in_word", []) + hierarchy_res.get("hierarchy_mismatched", [])
    if combined_h:
        seen_types = set()
        for item in combined_h:
            for d in [item.get("缺失简略描述"), item.get("层级不匹配简略描述"), item.get("简略描述")]:
                if d and d != "-":
                    for part in d.split("\n"):
                        part = part.strip()
                        if part and part != "-":
                            prefix = part.split("：")[0] if "：" in part else part
                            if prefix not in seen_types:
                                points.append(part)
                                seen_types.add(prefix)

    # 4. 需求变更规模因子
    factors = results.get("factor_check", {})
    scale_val = factors.get("scale", {}).get("value")
    if scale_val == "预算":
        points.append("需求变更规模因子异常：当前为【预算】，标准应为【结算】，请确认是否正确。")

    # 5. 质量特性因子
    missing, minus_one_items, empty_items = [], [], []
    name_map = {"distributed": "分布式处理", "performance": "性能", "reliability": "可靠性",
                "multiple_sites": "多重站点"}
    for key, cn_name in name_map.items():
        f = factors.get(key, {})
        if not f or not f.get("found_in_text"): missing.append(cn_name); continue
        str_val = str(f.get("value", "")).strip()
        if str_val == "-1":
            minus_one_items.append(cn_name)
        elif str_val in ["", "None", "空", "无", "默认"]:
            empty_items.append(cn_name)

    if missing: points.append(f"质量及特征因子{'、'.join(missing)}因子缺少，请确认是否正确")
    if minus_one_items: points.append(f"质量及特征因子{'、'.join(minus_one_items)}为-1，请确认是否正确")
    if empty_items: points.append(f"质量及特征因子{'、'.join(empty_items)}为空，请确认是否正确")

    # 6. 功能过程匹配
    process_res = results.get("process_res", {})
    not_found_p = process_res.get("not_found_in_word", [])
    if not_found_p:
        names = [f"【{item.get('Excel功能点', '未知')}】" for item in not_found_p[:2]]
        points.append(
            f"拆分表功能过程 {'、'.join(names)}{'等' if len(not_found_p) > 2 else ''} 功能过程在需求规格书中未体现，建议功能过程逐个核对...")

    # 7. 数据移动类型
    move_res = results.get("move_res", {})
    if move_res and not move_res.get("is_valid"):
        failed_items = [i for i in move_res.get("items", []) if i["result"] != "合规"]
        if failed_items:
            err_samples = [f"【{i['process']}】({i['result']})" for i in failed_items[:2]]
            points.append(
                f"功能过程数据移动类型校验不合规：{'、'.join(err_samples)}{'等' if len(failed_items) > 2 else ''}{len(failed_items)}处不合规。标准：一个完整功能过程应以 E 开始，并以 W 或 X 结束。")

    # 8. 资产清单匹配
    asset_res = results.get("asset_res", {})
    if asset_res and not asset_res.get("skipped"):
        if not asset_res.get("is_valid"):
            points.append(f"资产清单匹配执行出错：{asset_res.get('error', '未知错误')}")
        else:
            asset_stats = asset_res.get("statistics", {})
            asset_total = asset_stats.get("总数", 0)
            mismatch_cnt = asset_stats.get("不匹配", 0)
            match_rate_str = asset_stats.get("匹配率", "0%")
            if asset_total == 0:
                points.append("资产清单匹配未执行：拆分表中未找到有效的层级数据。")
            elif mismatch_cnt > 0:
                missing_lines, mismatch_lines = [], []
                for item in asset_res.get("items", []):
                    if item.get("匹配状态") != "不匹配":
                        continue
                    desc = item.get("缺失简略描述", "")
                    if desc and desc != "无缺失" and desc not in missing_lines:
                        missing_lines.append(desc)
                    desc2 = item.get("层级不匹配简略描述", "")
                    if desc2 and desc2 != "无层级错位" and desc2 not in mismatch_lines:
                        mismatch_lines.append(desc2)
                issue_parts = []
                if missing_lines:
                    issue_parts.append("拆分表模块在资产清单无对应：" + "、".join(missing_lines[:2])
                                       + ("等" if len(missing_lines) > 2 else ""))
                if mismatch_lines:
                    issue_parts.append("拆分表与资产清单层级不一致：" + "、".join(mismatch_lines[:2])
                                       + ("等" if len(mismatch_lines) > 2 else ""))
                detail = ("；".join(issue_parts) + "。") if issue_parts else ""
                points.append(
                    f"资产清单匹配存在异常（匹配率{match_rate_str}，共{mismatch_cnt}处问题）。{detail}")

    # 9. 数据属性重复
    data_attr_res = results.get("data_attr_res", {})
    if data_attr_res and not data_attr_res.get("skipped"):
        if not data_attr_res.get("success"):
            points.append(f"数据属性重复检测执行出错：{data_attr_res.get('error', '未知错误')}")
        else:
            attr_stats = data_attr_res.get("statistics", {})
            total_marked = attr_stats.get("total_marked_rows", 0)
            if total_marked > 0:
                intra_dups = attr_stats.get("intra_duplicates", 0)
                cross_dups = attr_stats.get("cross_duplicates", 0)
                points.append(
                    f"数据属性重复检测发现 {total_marked} 行存在重复：单功能过程内重复 {intra_dups} 组，"
                    f"跨功能过程重复 {cross_dups} 组，请核对拆分表中的重复数据项。")

    if not points:
        html += "<div style='font-size: 15px; color: #10b981;'>✅ 所有校验项均正常通过。</div>"
    else:
        for i, p in enumerate(points, 1):
            html += f'<div style="margin-bottom: 25px; padding-left: 12px; border-left: 4px solid {border_color}; font-weight: 700;"><span style="font-size: 15px;">{i}. {p}</span></div>'

    html += "</div>"
    return html

