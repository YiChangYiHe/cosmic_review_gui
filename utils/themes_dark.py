"""深色模式主题样式"""

from extend.matcher_config import MatcherConfig


def get_dark_theme_stylesheet():
    config = MatcherConfig.load()
    t = config.get("theme", {})
    primary = t.get("primary_color", "#66ccff")
    font_family = t.get("font_family", "Microsoft YaHei UI")
    font_size = t.get("font_size", 14)

    # [P1] 新建任务类 CTA 按钮统一取 Design Token，三页保持一致
    from utils.themes import get_tokens as _get_tokens

    _cta = _get_tokens(True)

    bg_main = "#1f2937"
    bg_header = "#111827"
    text_color = t.get("text_color_dark", "#f3f4f6")
    border_color = "#374151"
    item_bg = "#111827"

    return f"""
    QMainWindow, QDialog {{
        background-color: {bg_main};
    }}
    QDialog#ReceiptFileDialog {{
        background-color: {bg_main};
    }}
    QWidget#CentralWidget {{
        background-color: {bg_main};
    }}
    QWidget#RightContainer, QStackedWidget#StackedWidget {{
        background-color: {bg_main};
    }}
    QWidget#HomePage {{
        background-color: transparent;
    }}
    QFrame#HeaderFrame {{
        background: {bg_header};
        border-bottom: 1px solid {border_color};
        border-bottom-left-radius: 15px;
        border-bottom-right-radius: 15px;
    }}
    QLabel {{
        color: {text_color};
        font-size: {font_size}px;
    }}
    QLabel[class="task-title"] {{
        color: #f9fafb;
        font-weight: bold;
    }}
    QLabel[class="task-meta"], QLabel#SectionDesc {{
        color: #9ca3af;
    }}
    QFrame#WelcomeFrame {{
        background-color: rgba(102, 204, 255, 0.1);
        border: 2px solid #1e293b;
        border-radius: 16px;
    }}
    QLabel[class="greeting-text"] {{
        color: #cbd5e1;
    }}
    QScrollArea {{
        background: transparent;
        border: none;
    }}
    QScrollArea > QWidget > QWidget {{
        background-color: {bg_main};
    }}
    QWidget#ScrollContent {{
        background: {bg_main};
    }}
    QFrame[class="TaskCard"] {{
        background: transparent;
        border: 1px solid {border_color};
    }}
    QFrame#BizGroup {{
        background-color: rgba(20, 33, 61, 0.8);
        border: 2px solid rgba(102, 204, 255, 0.5);
        border-radius: 16px;
    }}
    QFrame#CleanGroup {{
        background-color: rgba(20, 33, 61, 0.8);
        border: 2px solid rgba(102, 204, 255, 0.5);
        border-radius: 16px;
    }}
    QLabel#SectionTitle {{
        font-size: 18px;
        font-weight: bold;
        color: #ffffff;
    }}
    QFrame#UploadArea {{
        border: 3px dashed #4b5563;
        border-radius: 10px;
        background: transparent;
    }}
    QFrame#UploadArea:hover {{
        border-color: {primary};
        background: transparent;
    }}
    QFrame[class="FileQueueItem"] {{
        background-color: {item_bg};
        border: 1px solid {border_color};
        border-radius: 8px;
        padding: 12px;
        margin: 2px 0;
    }}
    QFrame[class="FileQueueItem"] QPushButton {{
        background: {border_color};
        border: 1px solid #4b5563;
        color: #9ca3af;
    }}
    QFrame[class="FileQueueItem"] QPushButton:hover {{
        background: #4c0519;
        color: #fb7185;
        border-color: #fb7185;
    }}
    QFrame[class="FileQueueItem"]:hover {{
        border-color: {primary};
        background-color: {bg_main};
    }}
    QFrame[class="TaskCard"] QLabel {{
        color: {text_color};
    }}
    QFrame[class="TaskCard"] #LogPanel {{
        background: transparent;
        color: #d1d5db;
        padding: 12px 16px;
        border-radius: 6px;
        border: 1px solid {border_color};
        border-left: 4px solid #059669;
        font-size: 13px;
    }}
    QLabel[class="step-label"] {{
        font-size: 14px;

        color: #d1d5db;
        padding: 4px 2px;
        border-radius: 4px;
        background: transparent;
    }}
    QLabel[class="step-label"][status="processing"] {{ color: #60a5fa; font-weight: bold; }}
    QLabel[class="step-label"][status="done"] {{ color: #34d399; font-weight: bold; }}
    QLabel[class="step-label"][status="finished"] {{ color: #93c5fd; font-weight: bold; }}
    QLabel[class="step-label"][status="error"] {{ color: #fb7185; font-weight: bold; }}

    /* ReReviewTaskCard 样式 - 深色模式 */
    QFrame#ReReviewTaskCard {{
        background-color: #0f172a;
        border: 1px solid #1e293b;
        border-radius: 12px;
        padding: 16px;
    }}
    QFrame#ReReviewTaskCard QLabel#ProjectTitle {{
        color: #f1f5f9;
        font-size: 18px;
        font-weight: bold;
    }}
    QFrame#ReReviewTaskCard QLabel#TimeLabel {{
        color: #cbd5e1;
        font-size: 12px;
    }}
    QFrame#ReReviewTaskCard QLabel#StatusLabel {{
        color: #cbd5e1;
        font-size: 13px;
    }}
    QFrame#ReReviewTaskCard QPushButton.FileBtn {{
        background-color: #1e293b;
        color: #f1f5f9;
        border: 1px solid #334155;
        border-radius: 6px;
        padding: 8px 16px;
        font-size: 13px;
        font-weight: 500;
    }}
    QFrame#ReReviewTaskCard QPushButton.FileBtn:hover {{
        background-color: #334155;
        border: 1px solid #60a5fa;
    }}

    QTableWidget {{
        background-color: {border_color};
        gridline-color: #4b5563;
        border: 1px solid #4b5563;
        color: {text_color};
    }}
    QTableWidget::item {{
        color: {text_color};
    }}
    QHeaderView::section {{
        background-color: #4b5563;
        color: #f9fafb;
        padding: 4px;
        border: 1px solid {border_color};
    }}
    QGroupBox {{
        font-weight: bold;
        color: #38bdf8;
        border: 2px solid #38bdf8;
        border-radius: 10px;
        margin-top: 12px;
        padding-top: 10px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 15px;
        top: 2px;
        padding: 0 5px;
        background-color: {bg_main};
    }}
    QLineEdit, QComboBox, QDateEdit, QSpinBox {{
        border: 1px solid #4b5563;
        border-radius: 6px;
        padding: 5px;
        background: {item_bg};
        color: {text_color};
        selection-background-color: {primary};
        selection-color: white;
    }}
    QDateEdit QLineEdit {{
        background: transparent;
        color: inherit;
        border: none;
        padding: 0;
    }}
    /* [P0] QComboBox / QDateEdit 箭头与弹层样式统一收敛到 utils.themes.control_qss，
       此处不再重复定义，避免深浅色切换时两套规则打架 */
    QPushButton {{
        border-radius: 6px;
        padding: 8px 16px;
        color: {text_color};
        background: {border_color};
        border: 1px solid #4b5563;
    }}
    /* 三个页面的新建任务按钮统一为同一种主按钮样式（取 Design Token） */
    QPushButton#UploadBtn, QPushButton#NewTaskBtn, QPushButton#ReceiptBtn {{
        background-color: {_cta['accent']};
        color: white;
        border: none;
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: bold;
    }}
    QPushButton#UploadBtn:hover, QPushButton#NewTaskBtn:hover, QPushButton#ReceiptBtn:hover {{
        background-color: {_cta['accent_hover']};
    }}
    QTextBrowser {{
        background-color: {item_bg};
        color: {text_color};
        border: 1px solid {border_color};
    }}
    QPushButton#ThemeBtn {{
        background: transparent;
        color: #cbd5e1;
        border: 1px solid #293548;
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
    }}
    QPushButton#ThemeBtn:hover {{
        background: #1e293b;
        border-color: #3b82f6;
    }}
    QScrollArea, QScrollArea > QWidget > QWidget {{
        background-color: {bg_main};
    }}
    QScrollArea {{
        border: none;
    }}

    QCheckBox, QRadioButton {{
        spacing: 12px;
        color: {text_color};
    }}
    QCheckBox::indicator, QRadioButton::indicator {{
        width: 22px;
        height: 22px;
        border: 2px solid #4b5563;
        border-radius: 6px;
        background: #000000;
    }}
    QCheckBox::indicator:hover {{
        border-color: {primary};
    }}
    QCheckBox::indicator:checked {{
        background: {primary};
        border-color: {primary};
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
    }}
    QRadioButton::indicator:checked {{
        background: #66ccff;
        border-color: #66ccff;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0Ij48Y2lyY2xlIGN4PSIxMiIgY3k9IjEyIiByPSI1IiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==);
    }}
    /* 日历窗口适配 */
    QCalendarWidget QWidget {{
        background-color: {bg_main};
        color: {text_color};
    }}
    QCalendarWidget QAbstractItemView:enabled {{
        background-color: {bg_main};
        color: {text_color};
        selection-background-color: {primary};
        selection-color: white;
    }}
    QCalendarWidget QAbstractItemView:disabled {{
        color: #4b5563;
    }}
    QCalendarWidget QMenu {{
        background-color: {bg_header};
        color: {text_color};
    }}
    QCalendarWidget QSpinBox {{
        background-color: {bg_header};
        color: {text_color};
        border: none;
    }}
    QCalendarWidget QToolButton {{
        background-color: transparent;
        color: {text_color};
    }}
    QFrame#Sidebar {{
        background-color: #111827 !important;
        border-right: 1px solid #374151;
    }}
    QFrame#Sidebar > QWidget {{
        background-color: #111827 !important;
    }}
    QPushButton#SidebarToggleBtn {{
        background: transparent;
        border: none;
        font-size: 20px;
        font-weight: bold;
        color: #9ca3af;
        border-radius: 5px;
    }}
    QPushButton#SidebarToggleBtn:hover {{
        background: #374151;
    }}
    QPushButton#SidebarThemeBtn {{
        background-color: rgba(102, 204, 255, 0.1);
        border: 2px solid #374151;
        border-radius: 12px;
        font-size: 24px;
        color: #66ccff;
        padding: 4px;
        transition: all 0.3s ease;
    }}
    QPushButton#SidebarThemeBtn:hover {{
        background-color: rgba(102, 204, 255, 0.2);
        border-color: #66ccff;
        transform: scale(1.1);
    }}
    QPushButton#SidebarThemeBtn:pressed {{
        background-color: rgba(102, 204, 255, 0.15);
        transform: scale(0.95);
    }}
    QFrame#SidebarItem {{
        background: transparent;
        border-left: 3px solid transparent;
        border-radius: 6px;
        margin: 0 10px;
    }}
    QFrame#SidebarItem:hover {{
        background: #1e293b;
    }}
    QFrame#SidebarItem[selected="true"] {{
        background: rgba(59, 130, 246, 0.18);
        border-left: 3px solid #3b82f6;
    }}
    /* 侧边栏文字 - 未选中状态 */
    QFrame#SidebarItem QLabel#SidebarItemText {{
        color: #cbd5e1 !important;
        font-size: 14px;
        font-weight: 600;
    }}
    QFrame#SidebarItem QLabel#SidebarItemIcon {{
        color: #f3f4f6 !important;
        font-size: 18px;
    }}
    /* 侧边栏文字 - 选中状态 */
    QFrame#SidebarItem[selected="true"] QLabel#SidebarItemText {{
        color: #f8fafc !important;
        font-weight: bold;
    }}
    QFrame#SidebarItem[selected="true"] QLabel#SidebarItemIcon {{
        color: #f8fafc !important;
    }}

    /* 主题选择按钮 (设置页面) */
    QPushButton#ThemeLightBtn, QPushButton#ThemeDarkBtn {{
        border: 1px solid #293548;
        border-radius: 8px;
        background: #111827;
        color: #9ca3af;
        padding: 5px;
    }}
    QPushButton#ThemeLightBtn:checked, QPushButton#ThemeDarkBtn:checked {{
        background: rgba(102, 204, 255, 0.2);
        color: #66ccff;
        font-weight: bold;
        border: 1px solid #66ccff;
    }}
    QPushButton#ThemeLightBtn:hover:!checked, QPushButton#ThemeDarkBtn:hover:!checked {{
        background: #374151;
        color: #f3f4f6;
    }}

    QFrame[class="SettingsSection"] {{
        background: transparent;
        border: none;
    }}
    QWidget#HomePage {{
        background: linear-gradient(135deg, #0f172a 0%, #1a1f3a 50%, #2d1b4e 100%);
    }}
    QFrame#WelcomeFrame {{
        background-color: rgba(102, 204, 255, 0.15);
        border: 2px solid rgba(102, 204, 255, 0.6);
        border-radius: 16px;
    }}
    QLabel[class="greeting-text"] {{
        color: #cbd5e1;
    }}
    QFrame#BizGroup {{
        background-color: rgba(20, 33, 61, 0.8);
        border: 2px solid rgba(102, 204, 255, 0.5);
        border-radius: 16px;
    }}
    QFrame#CleanGroup {{
        background-color: rgba(20, 33, 61, 0.8);
        border: 2px solid rgba(102, 204, 255, 0.5);
        border-radius: 16px;
    }}
    QWidget#HomePage QLabel {{
        border: none;
        background: transparent;
    }}
    QPushButton#PrimaryBtn {{
        background: {primary};
        color: white;
        border-radius: 6px;
        font-weight: bold;
    }}
    QPushButton#PrimaryBtn:hover {{
        background: #2f6fe0;
    }}

    /* 复选框增强 (深色) */
    QCheckBox {{
        spacing: 12px;
        color: {text_color};
    }}
    QCheckBox::indicator {{
        width: 22px;
        height: 22px;
        border: 2px solid #4b5563;
        border-radius: 6px;
        background: #000000;
    }}
    QCheckBox::indicator:hover {{
        border-color: #66ccff;
    }}
    QCheckBox::indicator:checked {{
        background: #66ccff;
        border-color: #66ccff;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
    }}

    QLabel[class="separator-line"] {{
        color: #4b5563;
        font-size: 18px;
        padding: 0px 10px;
        line-height: 1.2;
        min-width: 20px;
        max-width: 20px;
    }}

    QFrame[class="separator"] {{
        background-color: #374151;
        border: none;
        max-width: 1px;
        min-width: 1px;
    }}
    
    /* ContentSection - 内容区块样式 */
    QFrame[class="ContentSection"] {{
        background-color: {item_bg};
        border: 1px solid {border_color};
        border-radius: 12px;
    }}
    
    /* ModeBtn - 模式选择按钮样式 */
    QPushButton[class="ModeBtn"] {{
        background-color: {item_bg};
        color: {text_color};
        border: 2px solid {border_color};
        border-radius: 8px;
        padding: 8px 20px;
        font-size: 13px;
        font-weight: bold;
    }}
    QPushButton[class="ModeBtn"]:hover {{
        background-color: {border_color};
        border-color: {primary};
    }}
    QPushButton[class="ModeBtn"]:checked {{
        background-color: {primary};
        color: white;
        border-color: {primary};
    }}
    
    /* SecondaryBtn - 次要按钮样式 */
    QPushButton[class="SecondaryBtn"] {{
        background-color: {item_bg};
        color: {text_color};
        border: 1px solid {border_color};
        border-radius: 6px;
        padding: 8px 16px;
        font-size: 12px;
        font-weight: bold;
    }}
    QPushButton[class="SecondaryBtn"]:hover {{
        background-color: {border_color};
        border-color: {primary};
    }}
    QPushButton[class="SecondaryBtn"]:disabled {{
        background-color: {item_bg};
        color: #4b5563;
        border-color: #374151;
    }}
"""
