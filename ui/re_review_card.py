from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
    QProgressBar,
)
from PySide6.QtCore import Qt, QVariantAnimation, QEasingCurve
import os
import subprocess
from utils.app_logger import ApplicationLogger

# 卡片类型 → (左侧彩条/徽章用的 token 键, 类型徽章文字)
# 重评用 warn 橙、回单用 success 绿，与历史页类型色一致
_CARD_TYPE_META = {
    "re_review": ("warn", "重评"),
    "receipt": ("success", "回单"),
}

# 卡片内通用的淡底/描边都需要 rgba 形式，统一从主题模块取
from utils.themes import rgba as _rgba
from .task_card import ElidedTitleLabel
from utils.runtime_logger import log_error


class ReReviewTaskCard(QFrame):
    """重评/回单任务卡片（card_type 区分左侧彩条与类型徽章）"""

    def __init__(self, task_info, card_type="re_review", parent=None):
        super().__init__(parent)
        self.task_info = task_info
        self.card_type = card_type if card_type in _CARD_TYPE_META else "re_review"
        self._shown_progress = 0.0  # 进度条当前显示值（动画驱动）
        self._pill_state = None  # running / queued / done / error
        self._error_mode = False  # 错误态按钮配色需要跟随主题重刷
        self._colored_labels = []  # [(QLabel, token_key, font_size, weight)]
        self._sub_labels = []  # [(QLabel, font_size)]
        self.setup_ui()
        self.update_theme_style()  # 初始化样式

    def showEvent(self, event):
        """每次显示时重新检查样式，防止主题切换后旧卡片没变色"""
        super().showEvent(event)
        self.update_theme_style()

    # ------------------------------------------------------------------
    # 主题与样式
    # ------------------------------------------------------------------
    def _detect_dark(self):
        """从配置检测主题模式，配置未设置时兜底扫描全局 QSS"""
        from PySide6.QtWidgets import QApplication
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        if not is_dark:
            qss = QApplication.instance().styleSheet() or ""
            is_dark = "background-color: #1f2937" in qss
        return is_dark

    def update_theme_style(self):
        """动态更新主题样式（颜色统一取自 utils.themes Design Token）"""
        from utils.themes import get_tokens

        self._is_dark = self._detect_dark()
        t = get_tokens(self._is_dark)
        type_color = t[_CARD_TYPE_META[self.card_type][0]]

        self.setStyleSheet(
            f"""
            QFrame#ReReviewTaskCard {{
                background-color: {t['bg_card']};
                border: 1px solid {t['border']};
                border-left: 4px solid {type_color};
                border-radius: 16px;
                padding: 16px 20px;
            }}
            /* 标签与内部容器一律透明，防止全局 QSS 刷出底色色块 */
            QLabel {{
                background: transparent;
            }}
            QWidget#CardInnerBox {{
                background: transparent;
            }}
            QLabel#ProjectTitle {{
                font-size: 18px;
                font-weight: 700;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {t['text_hi']};
                background: transparent;
            }}
            QLabel#TimeLabel {{
                font-size: 12px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {t['text_sub']};
                background: transparent;
            }}
            QLabel#StatusLabel {{
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {t['text_sub']};
                background: transparent;
            }}
            QLabel#ProgressPercent {{
                font-size: 13px;
                font-weight: 700;
                color: {t['accent']};
                background: transparent;
            }}
            QFrame#StatsFrame {{
                background-color: {t['bg_card2']};
                border: 1px solid {t['border']};
                border-radius: 10px;
                padding: 4px;
            }}
            QFrame#StatsFrame QLabel {{
                background: transparent;
            }}
            QFrame[objectName="DividerFrame"] {{
                background-color: {t['border']};
                border: none;
            }}
            QProgressBar {{
                background-color: {t['bg_card2']};
                border: 1px solid {t['border']};
                border-radius: 4px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                border-radius: 3px;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 {t['accent']}, stop:1 {t['accent_light']}
                );
            }}
            QPushButton.FileBtn {{
                border-radius: 8px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: 500;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                background-color: {t['bg_hover'] if self._is_dark else t['bg_card']};
                color: {t['text_hi']};
                border: 1px solid {t['border']};
            }}
            QPushButton.FileBtn:hover {{
                background-color: {t['bg_hover']};
                border: 1px solid {t['accent']};
            }}
            QPushButton.FileBtn:pressed {{
                background-color: {t['bg_hover']};
            }}
            QPushButton#OpenDirBtn {{
                background-color: {t['accent']};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: 500;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
            }}
            QPushButton#OpenDirBtn:hover {{
                background-color: {t['accent_hover']};
            }}
            QPushButton#OpenDirBtn:pressed {{
                background-color: {t['accent_hover']};
            }}
        """
        )

        # 类型徽章（淡底 + 描边，跟随类型色）
        if hasattr(self, "type_chip"):
            self.type_chip.setStyleSheet(
                f"color: {type_color}; background: {_rgba(type_color, 0.12)}; "
                f"border: 1px solid {_rgba(type_color, 0.28)}; border-radius: 4px; "
                "padding: 2px 8px; font-size: 11px; font-weight: 700; background-clip: padding-box;"
            )

        self._apply_status_pill(getattr(self, "_pill_state", None) or "running")
        self._refresh_dynamic_colors()

        # 错误态的颜色也跟随主题重刷
        if self._error_mode:
            self._apply_error_colors()

    def _apply_status_pill(self, state):
        """状态胶囊：进行中(蓝) / 排队中(灰) / 已完成(绿) / 运行出错(红)"""
        from utils.themes import get_tokens

        self._pill_state = state
        if not hasattr(self, "status_pill"):
            return
        t = get_tokens(getattr(self, "_is_dark", False))
        preset = {
            "running": ("● 进行中", "accent"),
            "queued": ("● 排队中", "text_sub"),
            "done": ("● 已完成", "success"),
            "error": ("● 运行出错", "danger"),
        }.get(state, ("● 进行中", "accent"))
        color = t[preset[1]]
        self.status_pill.setText(preset[0])
        self.status_pill.setStyleSheet(
            f"color: {color}; background: {_rgba(color, 0.12)}; "
            f"border: 1px solid {_rgba(color, 0.25)}; border-radius: 9px; "
            "padding: 2px 10px; font-size: 12px; font-weight: 600; background-clip: padding-box;"
        )

    def _set_status_text(self, key, bold=True):
        """状态文字着色（success/danger），并记录以便主题切换后重刷"""
        from utils.themes import get_tokens

        self._status_color_key = key
        t = get_tokens(getattr(self, "_is_dark", False))
        if not hasattr(self, "log_label"):
            return
        try:
            self.log_label.setStyleSheet(
                f"color: {t[key]}; font-size: 13px; font-weight: {700 if bold else 500};"
            )
        except RuntimeError:
            pass

    def _apply_error_colors(self):
        """错误态按钮配色（淡红底 + 红字，深浅主题各自取 token）"""
        from utils.themes import get_tokens

        t = get_tokens(getattr(self, "_is_dark", False))
        if hasattr(self, "log_btn"):
            try:
                self.log_btn.setStyleSheet(
                    f"background-color: {_rgba(t['danger'], 0.12)}; "
                    f"color: {t['danger']}; border: 1px solid {_rgba(t['danger'], 0.3)};"
                )
            except RuntimeError:
                pass

    def _refresh_dynamic_colors(self):
        """主题切换时重刷所有动态着色的统计标签"""
        from utils.themes import get_tokens

        t = get_tokens(getattr(self, "_is_dark", False))
        color_map = {
            "success": t["success"],
            "accent": t["accent"],
            "muted": t["text_sub"],
            "hi": t["text_hi"],
            "danger": t["danger"],
            "warn": t["warn"],
        }
        for lbl, key, size, weight in getattr(self, "_colored_labels", []):
            try:
                lbl.setStyleSheet(
                    f"color: {color_map[key]}; font-size: {size}px; "
                    f"font-weight: {weight};"
                )
            except RuntimeError:
                pass  # 控件可能已被销毁（合并模式重建后），忽略
        for lbl, size in getattr(self, "_sub_labels", []):
            try:
                lbl.setStyleSheet(
                    f"color: {t['text_sub']}; font-size: {size}px; font-weight: 400;"
                )
            except RuntimeError:
                pass

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def setup_ui(self):
        self.setObjectName("ReReviewTaskCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # ── Header：类型徽章 + 标题 + 状态胶囊 + 时间 ──
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        self.type_chip = QLabel(_CARD_TYPE_META[self.card_type][1])
        self.type_chip.setObjectName("TypeChip")
        header_layout.addWidget(self.type_chip)

        # 省略号标题：窗口不够宽时自动收缩，避免把右侧状态/按钮挤出可视区
        self.title_label = ElidedTitleLabel(
            self.task_info.get("project_name", "未命名项目")
        )
        self.title_label.setObjectName("ProjectTitle")
        self.title_label.setProperty("class", "task-title")
        header_layout.addWidget(self.title_label)
        header_layout.addStretch()

        status_text = self.task_info.get("status_text", "准备就绪")
        self._pill_state = "queued" if "等待" in status_text else "running"
        self.status_pill = QLabel("● 进行中")
        self.status_pill.setObjectName("StatusPill")
        header_layout.addWidget(self.status_pill)

        time_label = QLabel(f"处理时间: {self.task_info.get('time', '--:--:--')}")
        time_label.setObjectName("TimeLabel")
        time_label.setProperty("class", "task-meta")
        header_layout.addWidget(time_label)
        layout.addLayout(header_layout)

        # ── 状态文字 ──
        self.log_label = QLabel(status_text)
        self.log_label.setObjectName("StatusLabel")
        self.log_label.setWordWrap(True)
        layout.addWidget(self.log_label)

        # ── 进度行：渐变进度条 + 百分比 ──
        self.progress_row = QWidget()
        self.progress_row.setObjectName("CardInnerBox")  # 保持透明
        progress_layout = QHBoxLayout(self.progress_row)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        progress_layout.setSpacing(10)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        progress_layout.addWidget(self.progress_bar, stretch=1)

        self.percent_label = QLabel("0%")
        self.percent_label.setObjectName("ProgressPercent")
        self.percent_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.percent_label.setFixedWidth(44)
        progress_layout.addWidget(self.percent_label)

        layout.addWidget(self.progress_row)

        # 进度平滑动画（约 300ms 缓动追赶目标值）
        self._progress_anim = QVariantAnimation(self)
        self._progress_anim.setDuration(300)
        self._progress_anim.setEasingCurve(QEasingCurve.OutCubic)
        self._progress_anim.valueChanged.connect(self._on_anim_progress)

        # ── 统计数据卡片区域（初始隐藏） ──
        self.stats_widget = self._create_stats_widget()
        self.stats_widget.hide()
        layout.addWidget(self.stats_widget)

        layout.addSpacing(2)

        # ── 文件按钮区（完成前隐藏） ──
        self.files_widget = QWidget()
        self.files_widget.setObjectName("CardInnerBox")  # 保持透明
        files_layout = QHBoxLayout(self.files_widget)
        files_layout.setContentsMargins(0, 0, 0, 0)
        files_layout.setSpacing(10)

        btn_excel1 = QPushButton("📊 评估报告 (公式版)")
        btn_excel1.setProperty("class", "FileBtn")
        btn_excel1.setCursor(Qt.PointingHandCursor)
        self.excel1_btn = btn_excel1  # Save for later update
        btn_excel1.clicked.connect(lambda: self.open_file(self.task_info.get("excel1")))

        btn_excel2 = QPushButton("📋 重评回单 (结果版)")
        btn_excel2.setProperty("class", "FileBtn")
        btn_excel2.setCursor(Qt.PointingHandCursor)
        self.excel2_btn = btn_excel2  # Save for later update
        btn_excel2.clicked.connect(lambda: self.open_file(self.task_info.get("excel2")))

        # 查看日志按钮
        self.log_btn = QPushButton("📜 查看处理日志")
        self.log_btn.setProperty("class", "FileBtn")
        self.log_btn.setCursor(Qt.PointingHandCursor)
        self.log_btn.clicked.connect(self.open_log)

        btn_dir = QPushButton("📂 打开结果目录")
        btn_dir.setObjectName("OpenDirBtn")
        btn_dir.setCursor(Qt.PointingHandCursor)
        btn_dir.clicked.connect(
            lambda: self.open_file(self.task_info.get("output_dir"))
        )

        files_layout.addWidget(btn_excel1)
        files_layout.addWidget(btn_excel2)
        files_layout.addWidget(self.log_btn)
        files_layout.addStretch()
        files_layout.addWidget(btn_dir)

        layout.addWidget(self.files_widget)
        self.files_widget.hide()  # Hide until complete

    def _make_metric_group(self, sub_text, color_key, compact=False):
        """创建一个指标组：数值（着色加粗）在上，说明（灰色小字）在下"""
        val_size = 14 if compact else 16
        sub_size = 10 if compact else 11

        val = QLabel("0")
        self._colored_labels.append((val, color_key, val_size, 700))
        sub = QLabel(sub_text)
        self._sub_labels.append((sub, sub_size))

        box = QVBoxLayout()
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(2)
        box.addWidget(val)
        box.addWidget(sub)

        container = QWidget()
        container.setObjectName("CardInnerBox")  # 保持透明
        container.setLayout(box)
        return {"widget": container, "value": val, "sub": sub}

    def _create_stats_widget(self):
        """创建统计数据显示卡片（指标格风格，与初评卡片指标行同语言）"""
        widget = QWidget()
        widget.setObjectName("CardInnerBox")  # 保持透明
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        frame = QFrame()
        frame.setObjectName("StatsFrame")
        grid = QVBoxLayout(frame)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setSpacing(10)

        # 第一行：新增 / 复用 / 利旧 / 合计
        fp_row = QHBoxLayout()
        fp_row.setSpacing(28)
        self.g_new = self._make_metric_group("新增", "success")
        self.g_reuse = self._make_metric_group("复用", "accent")
        self.g_legacy = self._make_metric_group("利旧", "muted")
        self.g_total = self._make_metric_group("合计", "hi")
        for g in (self.g_new, self.g_reuse, self.g_legacy, self.g_total):
            fp_row.addWidget(g["widget"])
        fp_row.addStretch()
        grid.addLayout(fp_row)

        # 分隔线
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setLineWidth(1)
        divider.setObjectName("DividerFrame")
        grid.addWidget(divider)

        # 第二行：送审人天 / 核定人天 / 核减比例
        days_row = QHBoxLayout()
        days_row.setSpacing(28)
        self.g_submission = self._make_metric_group("送审人天", "hi")
        self.g_eval = self._make_metric_group("核定人天", "hi")
        self.g_reduction = self._make_metric_group("核减比例", "danger")
        for g in (self.g_submission, self.g_eval, self.g_reduction):
            days_row.addWidget(g["widget"])
        days_row.addStretch()
        grid.addLayout(days_row)

        layout.addWidget(frame)
        return widget

    # ------------------------------------------------------------------
    # 数据更新
    # ------------------------------------------------------------------
    def update_progress(self, value):
        """更新进度条（动画平滑过渡到目标值）"""
        try:
            target = max(0, min(100, int(value)))
        except (TypeError, ValueError):
            return
        # 已完成/出错后忽略迟到的进度信号，避免状态被改回"进行中"
        if self._pill_state in ("done", "error"):
            return
        self._apply_status_pill("running")
        if hasattr(self, "log_label"):
            try:
                if not self.log_label.text().startswith("正在处理"):
                    self.log_label.setText("正在处理，请稍候...")
            except RuntimeError:
                pass
        self._progress_anim.stop()
        self._progress_anim.setStartValue(float(self._shown_progress))
        self._progress_anim.setEndValue(float(target))
        self._progress_anim.start()

    def _on_anim_progress(self, val):
        self._shown_progress = float(val)
        v = int(round(val))
        if hasattr(self, "progress_bar"):
            self.progress_bar.setValue(v)
        if hasattr(self, "percent_label"):
            self.percent_label.setText(f"{v}%")

    def update_stats(self, stats):
        """更新统计数据显示"""
        if not stats:
            return

        # 如果 stats 是列表，说明是合并模式
        if isinstance(stats, list):
            # 清空现有统计显示重新构建（旧引用一并作废）
            self._colored_labels = []
            self._sub_labels = []
            layout = self.stats_widget.layout()
            while layout.count():
                item = layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            # 为每个子项目添加统计显示
            for sub_stat in stats:
                sub_name = sub_stat.get(
                    "sub_project_name", "合计" if sub_stat.get("is_total") else "子项目"
                )
                frame = self._create_mini_stats_frame(sub_stat, sub_name)
                layout.addWidget(frame)

            # 显示统计卡片
            self.stats_widget.show()
            return

        # 安全地获取值，处理N/A和转换类型
        def safe_get(val, default="N/A"):
            if val == "N/A" or val is None:
                return default
            return str(val)

        new_fp = safe_get(stats.get("new_fp", "N/A"), "0")
        reuse_fp = safe_get(stats.get("reuse_fp", "N/A"), "0")
        legacy_fp = safe_get(stats.get("legacy_fp", "N/A"), "0")
        total_fp = safe_get(stats.get("total_fp", "N/A"), "0")

        new_ratio = safe_get(stats.get("new_ratio", "0.0%"), "0.0%")
        reuse_ratio = safe_get(stats.get("reuse_ratio", "0.0%"), "0.0%")
        legacy_ratio = safe_get(stats.get("legacy_ratio", "0.0%"), "0.0%")
        total_ratio = safe_get(stats.get("total_ratio", "0.0%"), "0.0%")

        submission_days = safe_get(stats.get("submission_days", "0.0000"), "0.00")
        if "." in submission_days:
            try:
                submission_days = f"{float(submission_days):.2f}"
            except:
                pass

        eval_days = safe_get(stats.get("eval_days", "0.00"), "0.00")
        if "." in eval_days:
            try:
                eval_days = f"{float(eval_days):.2f}"
            except:
                pass

        reduction_ratio = safe_get(stats.get("reduction_ratio", "0.00%"), "0.00%")

        # 更新指标组（副标题保留指标名，避免丢失"新增/复用/利旧/合计"文字）
        self.g_new["value"].setText(new_fp)
        self.g_new["sub"].setText(f"新增 · 占比 {new_ratio}")
        self.g_reuse["value"].setText(reuse_fp)
        self.g_reuse["sub"].setText(f"复用 · 占比 {reuse_ratio}")
        self.g_legacy["value"].setText(legacy_fp)
        self.g_legacy["sub"].setText(f"利旧 · 占比 {legacy_ratio}")
        self.g_total["value"].setText(total_fp)
        self.g_total["sub"].setText(f"合计 · 占比 {total_ratio}")

        self.g_submission["value"].setText(submission_days)
        self.g_eval["value"].setText(eval_days)
        self.g_reduction["value"].setText(reduction_ratio)

        # 显示统计卡片
        self.stats_widget.show()

    def _create_mini_stats_frame(self, stats, title):
        """为合并模式创建紧凑的统计样式"""
        frame = QFrame()
        frame.setObjectName("StatsFrame")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(8)

        # 标题
        title_label = QLabel(f"📍 {title}")
        self._colored_labels.append((title_label, "hi", 13, 700))
        layout.addWidget(title_label)

        # 数据行1: 核心FP
        fp_row = QHBoxLayout()
        fp_row.setSpacing(20)
        for label, key, fp_k, ratio_k in (
            ("新增", "success", "new_fp", "new_ratio"),
            ("复用", "accent", "reuse_fp", "reuse_ratio"),
            ("利旧", "muted", "legacy_fp", "legacy_ratio"),
            ("合计", "hi", "total_fp", "total_ratio"),
        ):
            g = self._make_metric_group(
                f"{label} · {stats.get(ratio_k, '0.0%')}", key, compact=True
            )
            g["value"].setText(str(stats.get(fp_k, "0")))
            fp_row.addWidget(g["widget"])
        fp_row.addStretch()
        layout.addLayout(fp_row)

        # 数据行2: 人天
        try:
            sub_days = f"{float(stats.get('submission_days', 0)):.2f}"
        except:
            sub_days = stats.get("submission_days", "0")

        try:
            eval_days = f"{float(stats.get('eval_days', 0)):.2f}"
        except:
            eval_days = stats.get("eval_days", "0")

        days_row = QHBoxLayout()
        days_row.setSpacing(20)
        for value, label, key in (
            (sub_days, "送审人天", "hi"),
            (eval_days, "核定人天", "hi"),
            (stats.get("reduction_ratio", "0%"), "核减比例", "danger"),
        ):
            g = self._make_metric_group(label, key, compact=True)
            g["value"].setText(str(value))
            days_row.addWidget(g["widget"])
        days_row.addStretch()
        layout.addLayout(days_row)

        # 立即应用一次配色（主题切换后由 _refresh_dynamic_colors 重刷）
        self._refresh_dynamic_colors()
        return frame

    # ------------------------------------------------------------------
    # 状态切换
    # ------------------------------------------------------------------
    def set_completed(self, excel1=None, excel2=None):
        # 【新增】记录任务完成
        project_name = self.task_info.get("project_name", "未知项目")
        card_type = "重评" if self.card_type == "re_review" else "回单"
        ApplicationLogger.log_task_end(card_type, project_name, True)
        self._progress_anim.stop()
        self.progress_row.hide()
        self._error_mode = False
        self.log_label.show()
        self.log_label.setText("✅ 处理完成")
        self._set_status_text("success")
        self._apply_status_pill("done")
        self.files_widget.show()

        # 【修复】根据卡片类型设置不同的按钮文字
        if self.card_type == "re_review":
            # 重评场景：
            # excel1 = 新生成的重评报告
            # excel2 = 上一次的旧报告（供人工对比）
            if excel1:
                self.task_info["excel1"] = excel1
                self.excel1_btn.setText(" 查看历史报告 (旧)")
                try:
                    self.excel1_btn.clicked.disconnect()
                except:
                    pass
                self.excel1_btn.clicked.connect(lambda: self.open_file(excel1))
                self.excel1_btn.show()

            if excel2:
                self.task_info["excel2"] = excel2
                self.excel2_btn.setText("📊 查看重评报告 (新)")
                try:
                    self.excel2_btn.clicked.disconnect()
                except:
                    pass
                self.excel2_btn.clicked.connect(lambda: self.open_file(excel2))
                self.excel2_btn.show()

        elif self.card_type == "receipt":
            # 回单场景：
            # excel1 = 评估报告（已回写COSMIC结果）
            # excel2 = 评估确认单（Word文档）
            if excel1:
                self.task_info["excel1"] = excel1
                self.excel1_btn.setText(" 查看评估报告（原始）")
                try:
                    self.excel1_btn.clicked.disconnect()
                except:
                    pass
                self.excel1_btn.clicked.connect(lambda: self.open_file(excel1))
                self.excel1_btn.show()

            if excel2:
                self.task_info["excel2"] = excel2
                self.excel2_btn.setText("📋 查看评估确认单")
                try:
                    self.excel2_btn.clicked.disconnect()
                except:
                    pass
                self.excel2_btn.clicked.connect(lambda: self.open_file(excel2))
                self.excel2_btn.show()

        # 【新增】触发完成通知
        self._notify_finished(False)

    def set_error(self, message):
        """设置错误状态"""
        # 【新增】记录任务失败
        project_name = self.task_info.get("project_name", "未知项目")
        card_type = "重评" if self.card_type == "re_review" else "回单"
        ApplicationLogger.log_task_end(card_type, project_name, False, str(message))
        self._progress_anim.stop()
        self.progress_row.hide()
        self._error_mode = True
        self.log_label.show()
        # 限制错误消息长度，避免撑破布局
        short_msg = str(message)[:100] + ("..." if len(str(message)) > 100 else "")
        self.log_label.setText(f"❌ 运行报错: {short_msg}")
        self._set_status_text("danger")

        self._apply_status_pill("error")

        # 即使报错也显示按钮区域，以便查看日志
        self.files_widget.show()
        self.excel1_btn.hide()
        self.excel2_btn.hide()
        if hasattr(self, "log_btn"):
            self.log_btn.show()
            self.log_btn.setText("📜 查看错误日志")
            self._apply_error_colors()

        # 【修复】只调用一次错误通知
        self._notify_finished(True, str(message))

    # ------------------------------------------------------------------
    # 通知功能 (新增)
    # ------------------------------------------------------------------
    def _notify_finished(self, has_issue, error_msg=None):
        """任务完成或失败时的统一通知入口"""
        try:
            from extend.matcher_config import MatcherConfig
            from PySide6.QtWidgets import QApplication, QDialog, QPushButton, QSystemTrayIcon
            from PySide6.QtGui import QIcon
            from PySide6.QtCore import Qt, QTimer
            from utils.path_utils import get_resource_path

            config = MatcherConfig.load()
            auto_cfg = config.get("automation", {})
            flags = auto_cfg.get("notify_flags", {})

            # 兼容旧配置
            if not flags:
                mode = auto_cfg.get("notify_mode", "popup")
                flags = {"in_app": False, "popup": mode == "popup", "tray": mode == "tray"}

            # 如果三个都没选，直接返回
            if not any(flags.get(k) for k in ("in_app", "popup", "tray")):
                return

            win = self.window()
            project_name = self.task_info.get("project_name", "未知项目")

            # 根据卡片类型和状态生成标题
            card_type_name = "重评" if self.card_type == "re_review" else "回单"
            if has_issue:
                title = f"❌ {card_type_name}任务失败"
                msg = f"《{project_name}》处理出错：{error_msg or '未知错误'}"
            else:
                title = f"✅ {card_type_name}任务完成"
                msg = f"《{project_name}》处理完成，请查看结果。"

            # ================= 1. 任务栏闪烁 (独立逻辑) =================
            if win and not win.isActiveWindow():
                QApplication.alert(win, 2000)

            # ================= 2. 桌面通知 (Windows 右下角) =================
            if flags.get("tray"):
                try:
                    if getattr(self, "_tray", None) is None:
                        # 强制加载 logo.ico，解决显示 "C" 的问题
                        icon_path = get_resource_path("ui/logo.ico")
                        if not os.path.exists(icon_path):
                            icon_path = get_resource_path("ui/logo.png")

                        tray_icon = QIcon(icon_path) if os.path.exists(icon_path) else None
                        if tray_icon and not tray_icon.isNull():
                            self._tray = QSystemTrayIcon(tray_icon, win)
                        else:
                            self._tray = QSystemTrayIcon(win.windowIcon(), win)

                        self._tray.setToolTip("COSMIC 智能评估")
                        self._tray.show()

                    msg_type = QSystemTrayIcon.Warning if has_issue else QSystemTrayIcon.Information
                    self._tray.showMessage(title, msg, msg_type, 6000)
                except Exception as e:
                    log_error(f'[Tray Error] {e}')

            # ================= 3. 弹框通知 (中间的大窗口) =================
            if flags.get("popup"):
                try:
                    if getattr(self, "_notify_dialog", None) is not None:
                        try:
                            self._notify_dialog.close()
                        except:
                            pass

                    dlg = QDialog(win)
                    dlg.setWindowTitle("任务完成提醒")

                    # 【关键】给弹窗设置图标，解决标题显示 "C" 的问题
                    icon_path = get_resource_path("ui/logo.ico")
                    if os.path.exists(icon_path):
                        dlg.setWindowIcon(QIcon(icon_path))

                    dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowStaysOnTopHint)
                    dlg.setModal(False)
                    dlg.resize(420, 150)

                    from PySide6.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel
                    lay = QVBoxLayout(dlg)
                    title_label = QLabel(title)
                    title_label.setStyleSheet("font-size: 15px; font-weight: 700; color: #2563eb;")
                    lay.addWidget(title_label)

                    name_label = QLabel(msg)
                    name_label.setWordWrap(True)
                    lay.addWidget(name_label)

                    btn_row = QHBoxLayout()
                    btn_row.addStretch()
                    btn_close = QPushButton("关闭")
                    btn_row.addWidget(btn_close)
                    lay.addLayout(btn_row)

                    btn_close.clicked.connect(dlg.close)
                    QTimer.singleShot(15000, dlg.close)

                    self._notify_dialog = dlg
                    dlg.show()
                except Exception as e:
                    log_error(f'[Popup Error] {e}')

        except Exception as e:
            log_error(f'[Notify Error] {e}')

    # ------------------------------------------------------------------
    # 工具
    # ------------------------------------------------------------------
    def open_log(self):
        """打开对应的日志文件"""
        from utils.runtime_logger import RuntimeLogger

        log_path = RuntimeLogger.get_current_log_path(
            (self.task_info or {}).get("project_name")
        )
        if log_path:
            self.open_file(log_path)

    def open_file(self, path):
        if not path:
            return

        # 如果是列表，打开文件夹并高亮第一个（或只是打开文件夹）
        if isinstance(path, list):
            if not path:
                return
            target_path = os.path.dirname(path[0])
        else:
            target_path = path

        if not os.path.exists(target_path):
            return

        try:
            if os.name == "nt":
                if isinstance(path, list):
                    # 打开文件夹
                    os.startfile(target_path)
                else:
                    os.startfile(target_path)
            else:
                subprocess.call(
                    ["open" if os.name == "posix" else "xdg-open", target_path]
                )
        except Exception as e:
            log_error(f'无法打开路径: {target_path}, 错误: {e}')

    def update_title(self, title):
        """更新卡片标题"""
        if hasattr(self, "title_label"):
            try:
                self.title_label.setText(title)
                return
            except RuntimeError:
                pass

        label = self.findChild(QLabel, "ProjectTitle")
        if label:
            label.setText(title)
