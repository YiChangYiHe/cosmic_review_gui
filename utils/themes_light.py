"""浅色模式主题样式"""

from extend.matcher_config import MatcherConfig


def get_light_theme_stylesheet():
    config = MatcherConfig.load()
    t = config.get("theme", {})
    primary = t.get("primary_color", "#66ccff")
    font_family = t.get("font_family", "Microsoft YaHei UI")
    font_size = t.get("font_size", 14)

    # [P1] 新建任务类 CTA 按钮统一取 Design Token，三页保持一致
    from utils.themes import get_tokens as _get_tokens

    _tok = _get_tokens(False)

    bg_main = "#ffffff"
    bg_header = "#ffffff"
    text_color = "#000000"
    border_color = "#f1f5f9"
    item_bg = "#ffffff"

    return f"""
    QMainWindow, QDialog {{
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
        border-bottom: 2px solid #f8fafc;
    }}
    QLabel {{
        color: {text_color};
        font-size: {font_size}px;
    }}
    QLabel[class="task-title"] {{
        color: #000000;
        font-weight: bold;
    }}
    QLabel[class="task-meta"], QLabel#SectionDesc {{
        color: #000000;
    }}
    QLabel#WelcomeLabel {{
        color: #2563eb;
    }}
    QFrame#WelcomeFrame {{
        background-color: rgba(37, 99, 235, 0.12);
        border: 2px solid rgba(37, 99, 235, 0.6);
        border-radius: 16px;
    }}
    QLabel[class="greeting-text"] {{
        color: #64748b;
    }}
    /* 增加 ShortcutBtn 选择器支持 */
    QPushButton[class="ShortcutBtn"] {{
        border-radius: 8px;
        font-weight: 600;
        font-size: 14px;
    }}
    QScrollArea {{
        background: transparent;
        border: none;
    }}
    QScrollArea > QWidget > QWidget {{
        background-color: {bg_main};
    }}
    QWidget#ScrollContent {{
        background: transparent;
    }}
    QFrame[class="TaskCard"] {{
        background: white;
        border: 1px solid {border_color};
    }}
    QFrame#BizGroup, QFrame#CleanGroup {{
        background-color: rgba(255, 255, 255, 0.7);
        border: 2px solid rgba(37, 99, 235, 0.5);
        border-radius: 16px;
    }}
    QLabel#SectionTitle {{
        font-size: 18px;
        font-weight: bold;
        color: #1e293b;
    }}
    QFrame#UploadArea {{
        border: 3px dashed #cbd5e1;
        border-radius: 10px;
        background: white;
    }}
    QFrame#UploadArea:hover {{
        border-color: {primary};
        background: #f8fafc;
    }}
    QFrame[class="FileQueueItem"] {{
        background-color: white;
        border: 1px solid {border_color};
        border-radius: 8px;
        padding: 12px;
        margin: 2px 0;
    }}
    QFrame[class="FileQueueItem"] QPushButton {{
        background: #f1f5f9;
        border: 1px solid {border_color};
        color: #64748b;
    }}
    QFrame[class="FileQueueItem"] QPushButton:hover {{
        background: #fee2e2;
        color: #ef4444;
        border-color: #fca5a5;
    }}
    QFrame[class="FileQueueItem"]:hover {{
        border-color: {primary};
        background-color: #f8fafc;
    }}
    QLabel#LogPanel {{
        background: transparent;
        color: #000000;
        border-right: none;
        padding: 12px 16px;
        border-radius: 6px;
        border: 1px solid #cbd5e1;
        border-left: 4px solid #2563eb;
        font-size: 13px;
    }}
    QLabel[class="step-label"] {{
        font-size: 14px;

        color: #000000;
        padding: 4px 2px;
        border-radius: 4px;
    }}
    QLabel[class="step-label"][status="processing"] {{ color: {primary}; font-weight: bold; }}
    QLabel[class="step-label"][status="done"] {{ color: #10b981; font-weight: bold; }}
    QLabel[class="step-label"][status="finished"] {{ color: {primary}; font-weight: bold; }}
    QLabel[class="step-label"][status="error"] {{ color: #ef4444; font-weight: bold; }}

    /* ReReviewTaskCard 样式 - 浅色模式 */
    QFrame#ReReviewTaskCard {{
        background-color: #f0f4f8;
        border: 1px solid #cbd5e1;
        border-radius: 12px;
        padding: 16px;
    }}
    QFrame#ReReviewTaskCard QLabel#ProjectTitle {{
        color: #0f172a;
        font-size: 18px;
        font-weight: bold;
    }}
    QFrame#ReReviewTaskCard QLabel#TimeLabel {{
        color: #475569;
        font-size: 12px;
    }}
    QFrame#ReReviewTaskCard QLabel#StatusLabel {{
        color: #475569;
        font-size: 13px;
    }}
    QFrame#ReReviewTaskCard QPushButton.FileBtn {{
        background-color: #ffffff;
        color: #0f172a;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 8px 16px;
        font-size: 13px;
        font-weight: 500;
    }}
    QFrame#ReReviewTaskCard QPushButton.FileBtn:hover {{
        background-color: #f1f5f9;
        border: 1px solid #66ccff;
    }}

    QTableWidget {{
        background-color: white;
        gridline-color: {border_color};
        border: 1px solid {border_color};
        color: {text_color};
    }}
    QHeaderView::section {{
        background-color: #f1f5f9;
        color: #475569;
        padding: 4px;
        border: 1px solid {border_color};
    }}
    QGroupBox {{
        font-weight: bold;
        color: {primary};
        border: 2px solid {primary};
        border-radius: 10px;
        margin-top: 12px;
        padding-top: 10px;
    }}
    QLineEdit, QComboBox, QDateEdit, QSpinBox {{
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 5px;
        background: white;
        color: {text_color};
    }}
    QPushButton {{
        border-radius: 6px;
        padding: 8px 16px;
        color: {text_color};
        background: white;
        border: 1px solid #e2e8f0;
    }}
    /* 三个页面的新建任务按钮统一为同一种主按钮样式 */
    QPushButton#UploadBtn, QPushButton#NewTaskBtn, QPushButton#ReceiptBtn {{
        background-color: {_tok['accent']};
        color: white;
        border: none;
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: bold;
    }}
    QPushButton#UploadBtn:hover, QPushButton#NewTaskBtn:hover, QPushButton#ReceiptBtn:hover {{
        background-color: {_tok['accent_hover']};
    }}
    QTextBrowser {{
        background-color: white;
        color: {text_color};
        border: 1px solid {border_color};
    }}
    QPushButton#ThemeBtn {{
        background: transparent;
        color: #334155;
        border: 1px solid #d8dee9;
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
    }}
    QPushButton#ThemeBtn:hover {{
        background: #eef2f7;
        border-color: #2563eb;
    }}
    QFrame#Sidebar {{
        background-color: #e8ecf1 !important;
        border-right: 1px solid #d1dce6;
    }}
    QFrame#Sidebar > QWidget {{
        background-color: #e8ecf1 !important;
    }}
    QPushButton#SidebarToggleBtn {{
        background: transparent;
        border: none;
        font-size: 20px;
        font-weight: bold;
        color: #000000;
        border-radius: 5px;
    }}
    QPushButton#SidebarToggleBtn:hover {{
        background: #eff2f7;
    }}
    QPushButton#SidebarThemeBtn {{
        background: #f1f5f9;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        font-size: 24px;
        color: #0f172a;
        padding: 4px;
        transition: all 0.3s ease;
    }}
    QPushButton#SidebarThemeBtn:hover {{
        background: #dbeafe;
        border-color: #2563eb;
        transform: scale(1.1);
    }}
    QPushButton#SidebarThemeBtn:pressed {{
        background: #e0f2fe;
        transform: scale(0.95);
    }}
    /* [P0] QComboBox / QDateEdit 箭头与弹层样式统一收敛到 utils.themes.control_qss，
       此处不再重复定义，避免深浅色切换时两套规则打架 */
    /* 日历窗口适配 (亮色) */
    QCalendarWidget QWidget {{
        background-color: white;
        color: {text_color};
    }}
    QFrame#SidebarItem {{
        background: transparent;
        border-left: 3px solid transparent;
        border-radius: 6px;
        margin: 0 10px;
        padding-left: 0px;
    }}
    QFrame#SidebarItem:hover {{
        background: #eef2f7;
    }}
    QFrame#SidebarItem[selected="true"] {{
        background: #dbeafe;
        border-left: 3px solid #2563eb;
    }}
    /* 侧边栏文字 - 未选中状态 */
    QFrame#SidebarItem QLabel#SidebarItemText {{
        color: #334155 !important;
        font-size: 14px;
        font-weight: 600;
    }}
    QFrame#SidebarItem QLabel#SidebarItemIcon {{
        color: #334155 !important;
        font-size: 18px;
    }}
    /* 侧边栏文字 - 选中状态 ([FIX] 淡蓝底配深蓝字，白字看不清) */
    QFrame#SidebarItem[selected="true"] QLabel#SidebarItemText {{
        color: #1d4ed8 !important;
        font-weight: bold;
    }}
    QFrame#SidebarItem[selected="true"] QLabel#SidebarItemIcon {{
        color: #1d4ed8 !important;
    }}

    /* 主题选择按钮 (设置页面) */
    QPushButton#ThemeLightBtn, QPushButton#ThemeDarkBtn {{
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        background: transparent;
        color: #64748b;
        padding: 5px;
    }}
    QPushButton#ThemeLightBtn:checked, QPushButton#ThemeDarkBtn:checked {{
        background: rgba(37, 99, 235, 0.1);
        color: #2563eb;
        font-weight: bold;
        border: 1px solid #2563eb;
    }}
    QPushButton#ThemeLightBtn:hover:!checked, QPushButton#ThemeDarkBtn:hover:!checked {{
        background: #f1f5f9;
        color: #0f172a;
    }}

    QFrame[class="SettingsSection"] {{
        background: transparent;
        border: none;
    }}
    QPushButton#PrimaryBtn {{
        background: {primary};
        color: white;
        border-radius: 6px;
        font-weight: bold;
    }}
    QPushButton#PrimaryBtn:hover {{
        background: #1d4ed8;
    }}

    QCheckBox, QRadioButton {{
        spacing: 12px;
        color: {text_color};
    }}
    QCheckBox::indicator, QRadioButton::indicator {{
        width: 22px;
        height: 22px;
        border: 2px solid #cbd5e1;
        border-radius: 6px;
        background: #ffffff;
    }}
    QCheckBox::indicator:hover {{
        border-color: #66ccff;
    }}
    QCheckBox::indicator:checked {{
        background: #66ccff;
        border-color: #66ccff;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
    }}
    QRadioButton::indicator:checked {{
        background: #66ccff;
        border-color: #66ccff;
        image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0Ij48Y2lyY2xlIGN4PSIxMiIgY3k9IjEyIiByPSI1IiBmaWxsPSJ3aGl0ZSIvPjwvc3ZnPg==);
    }}

    QLabel[class="separator-line"] {{
        color: #cbd5e1;
        font-size: 18px;
        padding: 0px 10px;
        line-height: 1.2;
        min-width: 20px;
        max-width: 20px;
    }}

    QFrame[class="separator"] {{
        background-color: #e2e8f0;
        border: none;
        max-width: 1px;
        min-width: 1px;
    }}
    
    /* ContentSection - 内容区块样式 */
    QFrame[class="ContentSection"] {{
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
    }}
    
    /* ModeBtn - 模式选择按钮样式 */
    QPushButton[class="ModeBtn"] {{
        background-color: white;
        color: {text_color};
        border: 2px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 20px;
        font-size: 13px;
        font-weight: bold;
    }}
    QPushButton[class="ModeBtn"]:hover {{
        background-color: #f8fafc;
        border-color: {primary};
    }}
    QPushButton[class="ModeBtn"]:checked {{
        background-color: {primary};
        color: white;
        border-color: {primary};
    }}
    
    /* SecondaryBtn - 次要按钮样式 */
    QPushButton[class="SecondaryBtn"] {{
        background-color: white;
        color: {text_color};
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 8px 16px;
        font-size: 12px;
        font-weight: bold;
    }}
    QPushButton[class="SecondaryBtn"]:hover {{
        background-color: #f8fafc;
        border-color: {primary};
    }}
    QPushButton[class="SecondaryBtn"]:disabled {{
        background-color: #f8fafc;
        color: #cbd5e1;
        border-color: #e2e8f0;
    }}
"""
