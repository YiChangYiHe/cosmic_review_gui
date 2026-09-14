import sys
import os
import base64
from extend.matcher_config import MatcherConfig
from utils.themes_dark import get_dark_theme_stylesheet
from utils.themes_light import get_light_theme_stylesheet
from utils.runtime_logger import log_error


def apply_dark_title_bar(window, is_dark=True):
    """
    针对 Windows 10 (1809+) 和 Windows 11 开启原生标题栏深色模式
    """
    if sys.platform != "win32":
        return

    try:
        import ctypes
        from ctypes import wintypes

        # 获取 HWND，PySide6 的 winId() 返回的是个整数
        hwnd = int(window.winId())

        # DWMWA_USE_IMMERSIVE_DARK_MODE:
        # 20 是 Windows 11 及其以后版本
        # 19 是 Windows 10 (1809 - 21H1)
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19

        value = ctypes.c_int(1 if is_dark else 0)

        # 加载 dwmapi.dll
        dwmapi = ctypes.windll.dwmapi

        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE_OLD,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
    except Exception as e:
        log_error(f'Set dark title bar failed: {e}')



# =============================================================================
# [P0] Design Token：全 UI 统一的色彩/圆角规范，深浅两套主题各一份
# =============================================================================
_DESIGN_TOKENS = {
    False: {  # 浅色
        "bg_page": "#f4f6fa",
        "bg_card": "#ffffff",
        "bg_card2": "#f1f5f9",
        "bg_hover": "#eef2f7",
        "border": "#d8dee9",
        "text_hi": "#0f172a",
        "text_body": "#334155",
        "text_sub": "#64748b",
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "accent_light": "#60a5fa",
        "success": "#059669",
        "danger": "#dc2626",
        "warn": "#d97706",
        "input_bg": "#ffffff",
    },
    True: {  # 深色
        "bg_page": "#0b1220",
        "bg_card": "#111827",
        "bg_card2": "#0f172a",
        "bg_hover": "#1e293b",
        "border": "#293548",
        "text_hi": "#f8fafc",
        "text_body": "#cbd5e1",
        "text_sub": "#7c8ba1",
        "accent": "#3b82f6",
        "accent_hover": "#2f6fe0",
        "accent_light": "#93c5fd",
        "success": "#10b981",
        "danger": "#ef4444",
        "warn": "#f59e0b",
        "input_bg": "#0f172a",
    },
}


def get_tokens(is_dark=False):
    """获取设计 token（颜色规范）。所有组件样式应从这里取色。"""
    return _DESIGN_TOKENS[bool(is_dark)]


def is_dark_mode():
    """读取当前主题开关（单一事实来源，供各对话框替代"嗅探样式表"写法）"""
    try:
        return bool(MatcherConfig.load().get("theme", {}).get("is_dark", False))
    except Exception:
        return False


def _svg_b64(svg: str) -> str:
    """SVG 源码 → base64 data-uri，供 QSS image: 使用"""
    return base64.b64encode(svg.encode("utf-8")).decode("ascii")


def _chevron_svg(color: str, direction: str = "down") -> str:
    """生成主题色的下拉/上拉箭头 SVG"""
    points = "6 9 12 15 18 9" if direction == "down" else "6 15 12 9 18 15"
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" '
        f'fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" '
        f'stroke-linejoin="round"><polyline points="{points}"/></svg>'
    )


def rgba(hex_color, alpha):
    """#rrggbb → rgba(r, g, b, a) 字符串，用于淡底/描边"""
    h = (hex_color or "#000000").lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return f"rgba(0, 0, 0, {alpha})"
    return f"rgba({r}, {g}, {b}, {alpha})"


def control_qss(t):
    """
    全局基础控件样式（追加到主题样式表末尾）：
    下拉框（含弹层精装修）/输入框/复选框/滑块/工具提示/菜单/滚动条/禁用态。
    深浅色均从 Design Token 取色，切换主题即时生效。
    """
    chevron = _svg_b64(_chevron_svg(t["text_sub"], "down"))
    chevron_accent = _svg_b64(_chevron_svg(t["accent"], "down"))
    chevron_up = _svg_b64(_chevron_svg(t["text_sub"], "up"))
    accent_tint = rgba(t["accent"], 0.18)

    return f"""
    /* ===== [P0] 全局控件统一样式 ===== */
    QCheckBox {{
        color: {t['text_body']};
        spacing: 8px;
    }}
    QCheckBox::indicator {{
        width: 16px; height: 16px;
        border: 1px solid {t['border']};
        border-radius: 4px;
        background: {t['input_bg']};
    }}
    QCheckBox::indicator:hover {{
        border-color: {t['accent']};
    }}
    QCheckBox::indicator:checked {{
        background: {t['accent']};
        border-color: {t['accent']};
        image: url(data:image/svg+xml;base64,{_svg_b64('<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>')});
    }}

    /* ---- 下拉框（闭合态） ---- */
    QComboBox {{
        background: {t['input_bg']};
        color: {t['text_body']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 5px 12px;
        min-height: 22px;
    }}
    QComboBox:hover {{ border-color: {t['accent']}; }}
    QComboBox:focus {{ border-color: {t['accent']}; }}
    QComboBox:on {{ border-color: {t['accent']}; }}
    QComboBox:disabled {{ color: {t['text_sub']}; background: {t['bg_card2']}; }}
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: center right;
        border: none;
        width: 26px;
    }}
    QComboBox::down-arrow {{
        image: url(data:image/svg+xml;base64,{chevron});
        width: 14px; height: 14px;
    }}
    QComboBox:on::down-arrow {{
        image: url(data:image/svg+xml;base64,{chevron_accent});
    }}

    /* ---- 下拉弹层（QListView）：卡片化 + 圆角选中条目 + 无白箭头细滚动条 ---- */
    QComboBox QAbstractItemView {{
        background: {t['bg_card']};
        color: {t['text_body']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 4px;
        outline: none;
        selection-background-color: {accent_tint};
        selection-color: {t['text_hi']};
    }}
    /* 弹层容器（私有 QFrame，同为 QComboBox 子级）：不给底色会在深色弹层四周露出白边 */
    QComboBox QFrame {{
        background: {t['bg_card']};
        color: {t['text_body']};
        border: 1px solid {t['border']};
        border-radius: 8px;
    }}
    QComboBox QAbstractItemView::item {{
        min-height: 28px;
        padding: 4px 10px;
        margin: 1px 2px;
        border: none;
        border-radius: 6px;
        color: {t['text_body']};
        background: transparent;
    }}
    QComboBox QAbstractItemView::item:hover {{
        background: {accent_tint};
        color: {t['text_hi']};
    }}
    QComboBox QAbstractItemView::item:selected {{
        background: {accent_tint};
        color: {t['text_hi']};
    }}
    /* 弹层内滚动条：细、贴边、无箭头（杜绝浅色默认滚动条破坏深色弹层） */
    QComboBox QAbstractItemView QScrollBar:vertical {{
        background: transparent; width: 6px; margin: 2px; border: none;
    }}
    QComboBox QAbstractItemView QScrollBar::handle:vertical {{
        background: {t['border']}; border-radius: 3px; min-height: 24px;
    }}
    QComboBox QAbstractItemView QScrollBar::handle:vertical:hover {{ background: {t['text_sub']}; }}
    QComboBox QAbstractItemView QScrollBar::add-line:vertical,
    QComboBox QAbstractItemView QScrollBar::sub-line:vertical {{ height: 0px; width: 0px; }}
    QComboBox QAbstractItemView QScrollBar::add-page:vertical,
    QComboBox QAbstractItemView QScrollBar::sub-page:vertical {{ background: transparent; }}

    /* ---- 微调框（日期/数字）箭头同样跟随主题 ---- */
    QAbstractSpinBox::up-button, QAbstractSpinBox::down-button {{
        border: none; width: 20px; background: transparent;
    }}
    QAbstractSpinBox::up-button:hover, QAbstractSpinBox::down-button:hover {{
        background: {t['bg_hover']};
    }}
    QAbstractSpinBox::down-arrow {{
        image: url(data:image/svg+xml;base64,{chevron});
        width: 12px; height: 12px;
    }}
    QAbstractSpinBox::up-arrow {{
        image: url(data:image/svg+xml;base64,{chevron_up});
        width: 12px; height: 12px;
    }}

    QSlider::groove:horizontal {{
        height: 4px; border-radius: 2px;
        background: {t['border']};
    }}
    QSlider::sub-page:horizontal {{
        background: {t['accent']}; border-radius: 2px;
    }}
    QSlider::handle:horizontal {{
        width: 14px; height: 14px; margin: -6px 0;
        border-radius: 7px; background: {t['bg_card']};
        border: 2px solid {t['accent']};
    }}
    QSlider::handle:horizontal:hover {{ background: {t['accent']}; }}

    QToolTip {{
        background: {t['bg_card2']};
        color: {t['text_hi']};
        border: 1px solid {t['border']};
        padding: 5px 8px; border-radius: 4px;
    }}

    QMenu {{
        background: {t['bg_card']};
        color: {t['text_body']};
        border: 1px solid {t['border']};
        border-radius: 8px; padding: 4px;
    }}
    QMenu::item {{ padding: 6px 24px 6px 12px; border-radius: 4px; }}
    QMenu::item:selected {{ background: {t['accent']}; color: white; }}
    QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 8px; }}

    QScrollBar:vertical {{ background: transparent; width: 8px; margin: 2px; }}
    QScrollBar::handle:vertical {{
        background: {t['border']}; border-radius: 4px; min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{ background: {t['text_sub']}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; width: 0px; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    QScrollBar:horizontal {{ background: transparent; height: 8px; margin: 2px; }}
    QScrollBar::handle:horizontal {{
        background: {t['border']}; border-radius: 4px; min-width: 30px;
    }}
    QScrollBar::handle:horizontal:hover {{ background: {t['text_sub']}; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0px; height: 0px; }}
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: transparent; }}

    QLineEdit {{
        background: {t['input_bg']};
        color: {t['text_body']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 5px 10px;
    }}
    QLineEdit:focus {{ border-color: {t['accent']}; }}
    QPushButton:disabled {{ color: {t['text_sub']}; border-color: {t['border']}; }}
"""


def get_theme_stylesheet(is_dark=False):
    """根据主题类型返回对应的样式表"""
    tokens = get_tokens(is_dark)
    if is_dark:
        base = get_dark_theme_stylesheet()
    else:
        base = get_light_theme_stylesheet()
    # [P0] 追加全局基础控件样式（复选框/下拉/滑块/菜单/滚动条等）
    return base + "\n" + control_qss(tokens)


LIGHT_THEME = ""  # 为兼容性保留，实际改用 get_theme_stylesheet
DARK_THEME = ""
