# ui/evaluation_card.py
"""
评估任务卡片组件
显示COSMIC评估任务的状态和信息
"""

import os
from datetime import datetime
from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont


class EvaluationTaskCard(QFrame):
    """评估任务卡片"""

    # 信号
    view_clicked = Signal(str)  # 查看结果文件
    log_clicked = Signal()      # 查看日志文件
    delete_clicked = Signal(str)  # 删除任务

    def __init__(
        self,
        task_id,
        project_name,
        architecture_file,
        evaluation_file,
        output_file=None,
        log_path=None,
        status="待评估",
        create_time=None,
        parent=None,
    ):
        super().__init__(parent)
        self.task_id = task_id
        self.project_name = project_name
        self.architecture_file = architecture_file
        self.evaluation_file = evaluation_file
        self.output_file = output_file
        self.log_path = log_path
        self.status = status
        self.create_time = create_time or datetime.now()

        self.setObjectName("EvaluationTaskCard")
        self._setup_ui()

    def _setup_ui(self):
        self.setFrameShape(QFrame.StyledPanel)
        self.setFrameShadow(QFrame.Raised)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        # 顶部 - 项目名称和状态
        top_layout = QHBoxLayout()

        # 项目名称
        name_label = QLabel(self.project_name)
        name_label.setObjectName("ProjectName")
        name_font = QFont()
        name_font.setPointSize(12)
        name_font.setBold(True)
        name_label.setFont(name_font)
        top_layout.addWidget(name_label)

        top_layout.addStretch()

        # 状态标签
        self.status_label = QLabel(self.status)
        self.status_label.setObjectName("StatusLabel")
        self._update_status_style()
        top_layout.addWidget(self.status_label)

        layout.addLayout(top_layout)

        # 文件信息
        info_widget = QWidget()
        info_layout = QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 5, 0, 5)
        info_layout.setSpacing(5)

        # 功能架构图
        arch_label = QLabel(f"📄 功能架构图: {os.path.basename(self.architecture_file)}")
        arch_label.setObjectName("InfoLabel")
        info_layout.addWidget(arch_label)

        # 评估报告
        eval_label = QLabel(f"📊 评估报告: {os.path.basename(self.evaluation_file)}")
        eval_label.setObjectName("InfoLabel")
        info_layout.addWidget(eval_label)

        # 输出文件信息区域 (动态更新)
        self.output_info_widget = QWidget()
        self.output_info_layout = QVBoxLayout(self.output_info_widget)
        self.output_info_layout.setContentsMargins(0, 0, 0, 0)
        self.output_info_layout.setSpacing(5)
        
        self.output_label = QLabel("")
        self.output_label.setObjectName("InfoLabel")
        self.output_label.setVisible(False)
        self.output_info_layout.addWidget(self.output_label)
        
        info_layout.addWidget(self.output_info_widget)

        if self.output_file:
            self.set_output_file(self.output_file)

        # 创建时间
        time_str = self.create_time.strftime("%Y-%m-%d %H:%M")
        time_label = QLabel(f"🕐 创建时间: {time_str}")
        time_label.setObjectName("TimeLabel")
        info_layout.addWidget(time_label)

        layout.addWidget(info_widget)

        # 底部按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        # 查看结果按钮 (始终创建，动态控制显示)
        self.view_btn = QPushButton("查看结果")
        self.view_btn.setObjectName("ViewBtn")
        self.view_btn.setVisible(False)
        self.view_btn.clicked.connect(self._on_view_clicked)
        btn_layout.addWidget(self.view_btn)

        # 查看日志按钮
        self.log_btn = QPushButton("查看日志")
        self.log_btn.setObjectName("LogBtn")
        self.log_btn.setVisible(False)
        self.log_btn.clicked.connect(lambda: self.log_clicked.emit())
        btn_layout.addWidget(self.log_btn)

        if self.output_file:
            self.view_btn.setVisible(True)
            self.log_btn.setVisible(True)

        # 删除按钮
        self.delete_btn = QPushButton("删除")
        self.delete_btn.setObjectName("DeleteBtn")
        self.delete_btn.clicked.connect(lambda: self.delete_clicked.emit(self.task_id))
        btn_layout.addWidget(self.delete_btn)

        layout.addLayout(btn_layout)

        # 应用样式
        self._apply_styles()

    def _on_view_clicked(self):
        if self.output_file:
            self.view_clicked.emit(self.output_file)

    def _update_status_style(self):
        """更新状态标签样式"""
        status_colors = {
            "待评估": "#999",
            "评估中": "#0078d4",
            "已完成": "#107c10",
            "失败": "#d13438",
        }
        color = status_colors.get(self.status, "#999")
        self.status_label.setStyleSheet(
            f"""
            QLabel#StatusLabel {{
                background: {color};
                color: white;
                padding: 5px 15px;
                border-radius: 12px;
                font-weight: bold;
            }}
        """
        )

    def update_status(self, status):
        """更新状态"""
        self.status = status
        self.status_label.setText(status)
        self._update_status_style()

    def set_output_file(self, output_file):
        """设置输出文件并显示相应UI"""
        self.output_file = output_file
        if output_file:
            self.output_label.setText(f"✓ 输出文件: {os.path.basename(output_file)}")
            self.output_label.setVisible(True)
            self.view_btn.setVisible(True)
            self.log_btn.setVisible(True)

    def _apply_styles(self):
        """应用样式"""
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        
        if is_dark:
            bg_color = "#111827"
            border_color = "#374151"
            text_main = "#f3f4f6"
            text_sec = "#9ca3af"
            text_time = "#6b7280"
            btn_bg = "#3b82f6"
            btn_hover = "#2563eb"
        else:
            bg_color = "#ffffff"
            border_color = "#e2e8f0"
            text_main = "#1e293b"
            text_sec = "#64748b"
            text_time = "#94a3b8"
            btn_bg = "#2563eb"
            btn_hover = "#1d4ed8"

        self.setStyleSheet(
            f"""
            QFrame#EvaluationTaskCard {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 12px;
                margin: 5px;
            }}
            QFrame#EvaluationTaskCard:hover {{
                border-color: #3b82f6;
                background-color: {bg_color};
            }}
            QLabel#ProjectName {{
                color: {text_main};
            }}
            QLabel#InfoLabel {{
                color: {text_sec};
                font-size: 13px;
            }}
            QLabel#TimeLabel {{
                color: {text_time};
                font-size: 12px;
            }}
            QPushButton#ViewBtn {{
                background-color: {btn_bg};
                color: white;
                border: none;
                padding: 10px 24px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton#ViewBtn:hover {{
                background-color: {btn_hover};
            }}
            QPushButton#LogBtn {{
                background-color: #64748b;
                color: white;
                border: none;
                padding: 10px 24px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton#LogBtn:hover {{
                background-color: #475569;
            }}
            QPushButton#DeleteBtn {{
                background-color: transparent;
                color: #ef4444;
                border: 1px solid #fee2e2;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 13px;
            }}
            QPushButton#DeleteBtn:hover {{
                background-color: #fef2f2;
                border-color: #ef4444;
            }}
        """
        )
