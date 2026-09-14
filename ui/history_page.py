#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
历史项目页面：集中展示历史的初评 / 重评 / 回单项目记录。
支持同名项目合并折叠、状态显示、批量删除管理。
【终极修复版 v2】：补全遗漏的 _list_files 方法，确保多线程扫描 100% 稳定运行。
"""
import os
import re
import json
import shutil
from datetime import datetime, timedelta
from collections import defaultdict
from PySide6.QtCore import Qt, Signal, QThread, QTimer
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QScrollArea,
    QFrame, QFileDialog, QMessageBox, QLineEdit, QMenu, QCheckBox
)
from extend.matcher_config import MatcherConfig
from utils.task_state import is_heavy_task_running
from utils.themes import get_tokens
from .task_card import ElidedTitleLabel
from utils.runtime_logger import log_error, log_info, log_warn

_RUN_TS_RE = re.compile(r"(20\d{6})[_\-]?(\d{6})")
_TYPE_META = {
    "初评": {"icon": "📝", "color": "#3b82f6"},
    "重评": {"icon": "🔄", "color": "#f59e0b"},
    "回单": {"icon": "📄", "color": "#10b981"},
}
_GROUP_ORDER = ["今天", "昨天", "本周", "上周", "本月", "上月", "更早"]


def _parse_ts_from_name(name):
    m = _RUN_TS_RE.search(name or "")
    if m:
        try:
            datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
            return f"{m.group(1)}_{m.group(2)}"
        except ValueError:
            return None
    return None


def _mtime_ts(path):
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y%m%d_%H%M%S")
    except OSError:
        return "19700101_000000"


def _load_json_safe(path):
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


class _Storage:
    @staticmethod
    def _base(key):
        try:
            cfg = MatcherConfig.load()
            return os.path.abspath(cfg.get("storage", {}).get(key, "") or "")
        except Exception:
            return ""

    @classmethod
    def initial_review(cls):
        return cls._base("initial_review")

    @classmethod
    def re_review(cls):
        return cls._base("re_review")

    @classmethod
    def receipt(cls):
        return cls._base("receipt")

    @classmethod
    def logs(cls):
        return cls._base("logs")

# ==============================================================================
# 【核心修复】：后台扫描线程，彻底解放 UI 主线程，并在后台预计算状态
# ==============================================================================


class HistoryScanWorker(QThread):
    """后台扫描历史项目线程，彻底解放 UI 主线程"""
    finished = Signal(list)  # 扫描完成信号，传递 entries 列表

    def run(self):
        entries = []
        try:
            entries.extend(self._scan_initial_review())
            entries.extend(self._scan_re_review())
            entries.extend(self._scan_receipt())
            # 按时间戳倒序排序
            entries.sort(key=lambda e: e.get("ts", ""), reverse=True)
        except Exception as e:
            log_error(f'[ScanWorker] 扫描过程发生异常: {e}')
        self.finished.emit(entries)

    def _scan_initial_review(self):
        entries = []
        base = _Storage.initial_review()
        if not base or not os.path.isdir(base):
            return entries

        for project in self._list_dirs(base):
            p_dir = os.path.join(base, project)

            # 1. 先尝试查找时间戳子目录（新版结构）
            ts_dirs = []
            try:
                ts_dirs = [d for d in os.listdir(p_dir) if
                           os.path.isdir(os.path.join(p_dir, d)) and _parse_ts_from_name(d)]
            except OSError:
                continue

            if ts_dirs:
                # 新版结构：遍历时间戳子目录
                for d in ts_dirs:
                    ts = _parse_ts_from_name(d)
                    ts_dir = os.path.join(p_dir, d)

                    run_info = None
                    paused = None
                    log_file = None
                    summary_report = None
                    summary_html_path = None
                    count = 0
                    status_info = ("完成", "#10b981", None)
                    is_running = False

                    try:
                        files = os.listdir(ts_dir)
                    except OSError:
                        continue

                    for fn in files:
                        if fn.startswith("~$"):
                            continue
                        full_path = os.path.join(ts_dir, fn)

                        if fn == "run_info.json" or fn.endswith("-run_info.json"):
                            run_info = full_path
                            run_data = _load_json_safe(full_path)
                            if run_data:
                                # 【新增】检查是否正在运行
                                is_running = run_data.get("is_running", False)
                                summary_html_path = run_data.get("summary_html_path") or run_data.get("summary_html")
                        elif fn == "paused.json" or fn.endswith("-paused.json"):
                            paused = full_path
                        elif fn.endswith(".log"):
                            log_file = full_path
                        elif fn == "result.json":
                            res_data = _load_json_safe(full_path)
                            if res_data:
                                v_res = res_data.get("v_res", {})
                                summary_report = v_res.get("auto_report_path")
                                log_file = v_res.get("runtime_log_path") or log_file

                                passed, failed = self._count_status(v_res)
                                if failed > 0:
                                    status_info = ("异常", "#ef4444", f"正常{passed}，异常{failed}")
                                else:
                                    status_info = ("完成", "#10b981", f"全部{passed}项通过")
                        else:
                            count += 1

                    # 【新增】跳过正在运行的任务
                    if is_running:
                        continue

                    entries.append({
                        "type": "初评",
                        "project": project,
                        "ts": ts,
                        "time_text": self._fmt_ts(ts),
                        "detail": f"{count} 个产物文件",
                        "path": ts_dir,
                        "run_info": run_info,
                        "paused": paused,
                        "log_file": log_file,
                        "summary_report": summary_report,
                        "summary_html_path": summary_html_path,
                        "upload_manifest": run_info,
                        "status_info": status_info
                    })
            else:
                # 2. 旧版结构：没有时间戳子目录，使用 os.walk 遍历
                groups = {}
                try:
                    for root, _dirs, files in os.walk(p_dir):
                        for fn in files:
                            if fn.startswith("~$") or fn.endswith("-tree_cache.json"):
                                continue
                            full_path = os.path.join(root, fn)
                            ts = _parse_ts_from_name(fn) or _mtime_ts(full_path)

                            g = groups.setdefault(ts, {"count": 0, "run_info": None, "paused": None, "log_file": None,
                                                       "summary_html_path": None, "is_running": False})

                            if fn == "run_info.json" or fn.endswith("-run_info.json"):
                                g["run_info"] = full_path
                                run_data = _load_json_safe(full_path)
                                if run_data:
                                    g["is_running"] = run_data.get("is_running", False)
                                    g["summary_html_path"] = run_data.get("summary_html_path") or run_data.get(
                                        "summary_html")
                                continue
                            if fn == "paused.json" or fn.endswith("-paused.json"):
                                g["paused"] = full_path
                                continue
                            if fn.endswith(".log"):
                                g["log_file"] = full_path
                                continue
                            g["count"] += 1
                except OSError:
                    pass

                for ts in sorted(groups.keys(), reverse=True):
                    g = groups[ts]
                    # 【新增】跳过正在运行的任务
                    if g.get("is_running"):
                        continue

                    entries.append({
                        "type": "初评",
                        "project": project,
                        "ts": ts,
                        "time_text": self._fmt_ts(ts),
                        "detail": f"{g['count']} 个产物文件",
                        "path": p_dir,
                        "run_info": g["run_info"],
                        "paused": g["paused"],
                        "log_file": g.get("log_file"),
                        "summary_html_path": g.get("summary_html_path"),
                        "upload_manifest": g["run_info"],
                        "status_info": ("完成", "#10b981", None)
                    })
        return entries

    def _scan_re_review(self):
        entries = []
        base = _Storage.re_review()
        if not base or not os.path.isdir(base):
            return entries

        for name in self._list_dirs(base):
            d = os.path.join(base, name)

            # 1. 新版结构：{项目名}/{时间戳}/ 一次重评一个时间戳子目录。
            #    按子目录拆成多条记录：打开文件夹直达当次产物（最里层），文件数按次统计。
            ts_dirs = []
            try:
                ts_dirs = [x for x in os.listdir(d)
                           if os.path.isdir(os.path.join(d, x)) and _parse_ts_from_name(x)]
            except OSError:
                continue

            if ts_dirs:
                # 项目名优先取 inputs.json 里记录的原始项目名，保持与旧记录可合并
                project_name = ""
                for sub in ts_dirs:
                    m = _load_json_safe(os.path.join(d, sub, "inputs.json"))
                    if m and m.get("project_name"):
                        project_name = m["project_name"]
                        break
                project_name = project_name or name

                for sub in ts_dirs:
                    ts = _parse_ts_from_name(sub)
                    run_dir = os.path.join(d, sub)
                    manifest = os.path.join(run_dir, "inputs.json")
                    entries.append({
                        "type": "重评",
                        "project": project_name,
                        "ts": ts,
                        "time_text": self._fmt_ts(ts),
                        "detail": f"{self._count_files(run_dir)} 个文件",
                        "path": run_dir,
                        "upload_manifest": manifest if os.path.exists(manifest) else None,
                    })
                continue

            # 2. 旧版结构：顶层目录本身就是一次重评（{项目名}_重评_{时间戳} 等）
            ts = _parse_ts_from_name(name) or _mtime_ts(d)
            manifest = os.path.join(d, "inputs.json")

            # 【核心修复】：智能提取项目名称，兼容新旧版本
            project_name = ""

            # 1. 优先从 inputs.json 读取（新版）
            if os.path.exists(manifest):
                manifest_data = _load_json_safe(manifest)
                if manifest_data:
                    project_name = manifest_data.get("project_name", "")

            # 2. 旧版兼容：从文件夹名或内部文件名提取
            if not project_name:
                # 清理文件夹名前缀
                cleaned_name = re.sub(r"^重评结果[_\-]?", "", name).strip()

                # 如果文件夹名包含项目名（如 "重评结果_项目名_20260717_114647"）
                if cleaned_name and not re.match(r"^\d{8}_\d{6}$", cleaned_name):
                    project_name = cleaned_name
                else:
                    # 文件夹名只有时间戳，从内部文件提取
                    for fn in self._list_files(d):
                        # 找第一个 Excel 或 Word 文件
                        if fn.endswith((".xlsx", ".xls", ".docx", ".doc")) and not fn.startswith("~$"):
                            project_name = os.path.splitext(fn)[0]
                            # 清理文件名前缀（如 "xxxx-" 或 "x_"）
                            project_name = re.sub(r"^[xX_\-]+", "", project_name).strip()
                            if project_name:
                                break

                    # 如果还是没找到，用文件夹名兜底
                    if not project_name:
                        project_name = cleaned_name or name

            entries.append({
                "type": "重评",
                "project": project_name,
                "ts": ts,
                "time_text": self._fmt_ts(ts),
                "detail": f"{self._count_files(d)} 个文件",
                "path": d,
                "upload_manifest": manifest if os.path.exists(manifest) else None,
            })
        return entries

    def _scan_receipt(self):
        """【完整实现】扫描回单记录"""
        entries = []
        base = _Storage.receipt()
        if not base or not os.path.isdir(base):
            return entries
        # 1. 扫描文件夹形式的回单
        for name in self._list_dirs(base):
            d = os.path.join(base, name)
            ts = _parse_ts_from_name(name) or _mtime_ts(d)
            manifest = os.path.join(d, "inputs.json")
            entries.append({
                "type": "回单",
                "project": name,
                "ts": ts,
                "time_text": self._fmt_ts(ts),
                "detail": f"{self._count_files(d)} 个文件",
                "path": d,
                "upload_manifest": manifest if os.path.exists(manifest) else None,
            })
        # 2. 扫描散落在根目录的单个回单文件（兼容旧版结构）
        for fn in self._list_files(base):
            if fn.endswith("-inputs.json") or fn == "inputs.json" or fn.startswith("~$"):
                continue
            f = os.path.join(base, fn)
            ts = _parse_ts_from_name(fn) or _mtime_ts(f)
            stem = os.path.splitext(fn)[0]
            own = os.path.join(base, f"{stem}-inputs.json")
            manifest = own if os.path.exists(own) else None
            entries.append({
                "type": "回单",
                "project": stem,
                "ts": ts,
                "time_text": self._fmt_ts(ts),
                "detail": "单个文件",
                "path": base,
                "upload_manifest": manifest,
            })
        return entries

    # ================= 辅助方法（实例方法，非静态）=================

    @staticmethod
    def _list_dirs(base):
        try:
            return sorted([d for d in os.listdir(base) if
                           os.path.isdir(os.path.join(base, d)) and not d.startswith((".", "~$", "json_reports"))])
        except OSError:
            return []

    @staticmethod
    def _list_files(base):
        """【补全】列出目录下的有效文件，过滤临时文件和缓存"""
        try:
            return sorted([f for f in os.listdir(base) if
                           os.path.isfile(os.path.join(base, f)) and not f.startswith("~$") and not f.endswith(
                               "-tree_cache.json")])
        except OSError:
            return []

    @staticmethod
    def _count_files(path, recursive=True):
        """统计文件夹内的文件数量"""
        count = 0
        if recursive:
            for _root, _dirs, files in os.walk(path):
                count += len([f for f in files if not f.startswith("~$")])
        else:
            try:
                count = len([f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))])
            except OSError:
                count = 0
        return count

    @staticmethod
    def _count_status(v_res):
        """计算通过/失败数量"""
        passed = 0
        failed = 0
        checks = [
            ("excel_check", lambda r: r.get("is_ok", True)),
            ("ratio_check", lambda r: r.get("is_ok", True)),
            ("hierarchy_res", lambda r: r.get("statistics", {}).get("缺失项", 0) == 0),
            ("process_res", lambda r: r.get("statistics", {}).get("缺失项", 0) == 0),
            ("move_res", lambda r: r.get("statistics", {}).get("不合规", 0) == 0),
            ("asset_res", lambda r: r.get("statistics", {}).get("不匹配", 0) == 0),
            ("data_attr_res", lambda r: r.get("statistics", {}).get("total_marked_rows", 0) == 0),
        ]
        for key, is_ok_func in checks:
            res = v_res.get(key)
            if res and not res.get("skipped"):
                if is_ok_func(res):
                    passed += 1
                else:
                    failed += 1
        return passed, failed

    @staticmethod
    def _fmt_ts(ts):
        try:
            return datetime.strptime(ts, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M")
        except (ValueError, TypeError):
            return ts or "未知时间"

# ==============================================================================
# UI 组件
# ==============================================================================


class HistoryGroupCard(QFrame):
    """历史项目分组卡片：支持折叠/展开"""

    def __init__(self, project_name, entries, parent=None, batch_mode=False):
        super().__init__(parent)
        self.project_name = project_name
        self.entries = entries
        self.is_expanded = False
        self.parent_page = parent
        self.batch_mode = batch_mode
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.header = QFrame()
        self.header.setObjectName("GroupHeader")
        h_layout = QHBoxLayout(self.header)
        h_layout.setContentsMargins(16, 12, 16, 12)
        h_layout.setSpacing(12)

        if self.batch_mode:
            self.group_cb = QCheckBox()
            self.group_cb.setCursor(Qt.PointingHandCursor)
            self.group_cb.toggled.connect(self._on_group_toggled)
            h_layout.addWidget(self.group_cb)

        type_icon = QLabel(_TYPE_META.get(self.entries[0]["type"], {}).get("icon", ""))
        type_icon.setStyleSheet("font-size: 18px;")
        h_layout.addWidget(type_icon)

        etype = self.entries[0]["type"]
        meta = _TYPE_META.get(etype, {})
        badge = QLabel(etype)
        badge.setStyleSheet(
            f"background: {meta['color']}; color: white; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: 700;")
        badge.setFixedHeight(20)
        h_layout.addWidget(badge)

        name_lbl = ElidedTitleLabel(self.project_name)
        name_lbl.setStyleSheet("font-size: 15px; font-weight: 700;")
        h_layout.addWidget(name_lbl, 1)

        latest = self.entries[0]
        status_text, status_color, stats = latest.get("status_info", ("完成", "#10b981", None))
        if latest.get("paused"):
            paused_data = _load_json_safe(latest["paused"])
            if paused_data:
                status_text, status_color, stats = "暂停", "#f59e0b", f"进度 {paused_data.get('completed_through', 0)}/10"

        status_container = QWidget()
        status_layout = QHBoxLayout(status_container)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(8)

        if stats:
            stats_label = QLabel(stats)
            stats_label.setStyleSheet("color: #94a3b8; font-size: 12px; font-weight: 500; padding: 2px 6px;")
            status_layout.addWidget(stats_label)

        status_label = QLabel(status_text)
        status_label.setStyleSheet(
            f"color: {status_color}; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 12px; background: {self._get_soft_bg_color(status_color)}; border: 1px solid {self._get_border_color(status_color)};")
        status_layout.addWidget(status_label)
        h_layout.addWidget(status_container)

        count_lbl = QLabel(f"{len(self.entries)} 次")
        count_lbl.setStyleSheet("color: #64748b; font-size: 12px;")
        h_layout.addWidget(count_lbl)

        layout.addWidget(self.header)

        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(20, 10, 20, 10)
        self.content_layout.setSpacing(8)

        for e in self.entries:
            self.content_layout.addWidget(self._make_item_card(e))

        self.content_widget.setVisible(False)
        layout.addWidget(self.content_widget)

        self.header.setCursor(Qt.PointingHandCursor)
        self.header.mousePressEvent = lambda event: self._toggle()
        self._apply_style()

    def _on_group_toggled(self, checked):
        if not self.parent_page: return
        for e in self.entries:
            p = e.get("path")
            if p:
                if checked:
                    self.parent_page._selected_paths.add(p)
                else:
                    self.parent_page._selected_paths.discard(p)

    def _get_soft_bg_color(self, status_color):
        if status_color == "#ef4444":
            return "rgba(239, 68, 68, 0.08)"
        elif status_color == "#10b981":
            return "rgba(16, 185, 129, 0.08)"
        elif status_color == "#f59e0b":
            return "rgba(245, 158, 11, 0.08)"
        return "rgba(148, 163, 184, 0.08)"

    def _get_border_color(self, status_color):
        if status_color == "#ef4444":
            return "rgba(239, 68, 68, 0.2)"
        elif status_color == "#10b981":
            return "rgba(16, 185, 129, 0.2)"
        elif status_color == "#f59e0b":
            return "rgba(245, 158, 11, 0.2)"
        return "rgba(148, 163, 184, 0.2)"

    def _toggle(self):
        self.is_expanded = not self.is_expanded
        self.content_widget.setVisible(self.is_expanded)
        self._apply_style()

    def _apply_style(self):
        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)
        t = get_tokens(is_dark)
        border = t['accent'] if self.is_expanded else t['border']
        bg = "rgba(148,163,184,0.03)" if self.is_expanded else "transparent"
        self.header.setStyleSheet(
            f"QFrame#GroupHeader {{ background: {bg}; border: 1px solid {border}; border-radius: 8px; }} QFrame#GroupHeader:hover {{ background: {t['bg_hover']}; }}")

    def _make_item_card(self, e):
        meta = _TYPE_META[e["type"]]
        card = QFrame()
        card.setObjectName("ItemCard")
        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)
        t = get_tokens(is_dark)
        card.setStyleSheet(
            f"QFrame#ItemCard {{ background: {t['bg_card']}; border: 1px solid {t['border']}; border-left: 3px solid {meta['color']}; border-radius: 6px; }}")

        row = QHBoxLayout(card)
        row.setContentsMargins(12, 8, 12, 8)
        row.setSpacing(10)

        if self.batch_mode:
            cb = QCheckBox()
            cb.setCursor(Qt.PointingHandCursor)
            is_selected = e.get("path") in self.parent_page._selected_paths if self.parent_page else False
            cb.setChecked(is_selected)
            cb.toggled.connect(lambda checked, p=e.get("path"): self._on_item_selected(p, checked))
            row.addWidget(cb)

        time_lbl = QLabel(e['time_text'])
        time_lbl.setStyleSheet(f"color: {t['text_sub']}; font-size: 12px; min-width: 110px;")
        row.addWidget(time_lbl)

        detail_lbl = QLabel(e['detail'])
        detail_lbl.setStyleSheet(f"color: {t['text_body']}; font-size: 12px;")
        row.addWidget(detail_lbl, 1)

        if e.get("type") == "初评" and e.get("run_info"):
            btn_rerun = QPushButton("🔁 重评")
            btn_rerun.setFixedHeight(26)
            btn_rerun.setCursor(Qt.PointingHandCursor)
            btn_rerun.setStyleSheet(
                f"font-size: 11px; padding: 0 8px; border: 1px solid {t['border']}; border-radius: 4px; background: transparent; color: {t['text_body']};")
            if self.parent_page and hasattr(self.parent_page, '_emit_re_run'):
                btn_rerun.clicked.connect(lambda _=False, p=e.get("run_info"): self.parent_page._emit_re_run(p))
            row.addWidget(btn_rerun)

        btn_open = QPushButton(" 打开")
        btn_open.setFixedHeight(26)
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.setStyleSheet(
            f"font-size: 11px; padding: 0 8px; border: 1px solid {t['border']}; border-radius: 4px; background: transparent; color: {t['text_body']};")
        btn_open.clicked.connect(lambda _=False, p=e.get("path"): self._open_dir(p))
        row.addWidget(btn_open)

        card.setContextMenuPolicy(Qt.CustomContextMenu)
        if self.parent_page and hasattr(self.parent_page, '_show_card_context_menu'):
            card.customContextMenuRequested.connect(
                lambda pos, entry=e: self.parent_page._show_card_context_menu(pos, entry, card))
        return card

    def _on_item_selected(self, path, checked):
        if not self.parent_page: return
        if checked:
            self.parent_page._selected_paths.add(path)
        else:
            self.parent_page._selected_paths.discard(path)

    @staticmethod
    def _open_dir(path):
        try:
            if path and os.path.isdir(path):
                if os.name == "nt":
                    os.startfile(path)
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            log_error(f'打开文件夹失败: {exc}')


class HistoryPage(QWidget):
    """历史项目页"""
    re_run_requested = Signal(dict)
    resume_requested = Signal(dict)
    PAGE_SIZE = 20

    def __init__(self, parent=None):
        super().__init__(parent)
        self._entries = []
        self._filter = "全部"
        self._page = 1
        self._search_text = ""
        self._merge_enabled = MatcherConfig.load().get("history_display_mode", "merge") == "merge"
        self._batch_mode = False
        self._selected_paths = set()
        self._init_ui()

    def _is_dark(self):
        try:
            return MatcherConfig.load().get("theme", {}).get("is_dark", False)
        except:
            return False

    def _sub_color(self):
        return "#94a3b8" if self._is_dark() else "#64748b"

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QFrame()
        header.setObjectName("HistoryHeader")
        header.setFixedHeight(60)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)
        header_layout.setSpacing(15)

        title = QLabel("🕘 历史项目")
        title.setStyleSheet("font-size: 20px; font-weight: 800;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 搜索项目名...")
        self.search_edit.setFixedHeight(32)
        # 弹性宽度：固定 200px 会挤占右侧按钮，窄窗口下"批量删除"被截断
        self.search_edit.setMinimumWidth(130)
        self.search_edit.setMaximumWidth(210)
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._on_search_changed)
        header_layout.addWidget(self.search_edit, stretch=1)
        header_layout.addSpacing(10)

        self._filter_buttons = {}
        for key in ["全部", "初评", "重评", "回单"]:
            btn = QPushButton(key)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFixedHeight(32)
            btn.clicked.connect(lambda _=False, k=key: self._set_filter(k))
            self._filter_buttons[key] = btn
            header_layout.addWidget(btn)
        self._filter_buttons["全部"].setChecked(True)

        self._batch_btn = QPushButton("🗑️ 批量删除")
        self._batch_btn.setCheckable(True)
        self._batch_btn.setChecked(False)
        self._batch_btn.setFixedHeight(32)
        self._batch_btn.setCursor(Qt.PointingHandCursor)
        self._batch_btn.clicked.connect(self._toggle_batch_mode)
        header_layout.addWidget(self._batch_btn)

        self._refresh_btn = QPushButton(" 刷新")
        self._refresh_btn.setFixedHeight(32)
        self._refresh_btn.setCursor(Qt.PointingHandCursor)
        self._refresh_btn.clicked.connect(self.refresh)
        header_layout.addWidget(self._refresh_btn)

        layout.addWidget(header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(
            "QScrollArea { border: none; background: transparent; } QScrollBar:vertical { background: transparent; width: 8px; margin: 2px; } QScrollBar::handle:vertical { background: rgba(100,116,139,0.5); border-radius: 4px; min-height: 30px; } QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }")

        self.list_content = QWidget()
        self.list_content.setStyleSheet("background: transparent;")
        self.list_layout = QVBoxLayout(self.list_content)
        self.list_layout.setContentsMargins(20, 15, 20, 15)
        self.list_layout.setSpacing(12)
        self.list_layout.addStretch()

        scroll.setWidget(self.list_content)
        layout.addWidget(scroll)
        self._sync_filter_style()

    def _style_search_edit(self):
        """搜索框样式跟随主题（原先硬编码深色，浅色模式下白底看不清）"""
        t = get_tokens(self._is_dark())
        self.search_edit.setStyleSheet(
            f"QLineEdit {{ background: {t['input_bg']}; color: {t['text_hi']};"
            f" border: 1px solid {t['border']}; border-radius: 8px; padding: 0 10px; font-size: 13px; }}"
            f"QLineEdit:focus {{ border-color: {t['accent']}; }}"
        )

    def _style_batch_btn(self):
        """批量删除按钮样式跟随主题"""
        t = get_tokens(self._is_dark())
        if self._batch_btn.isChecked():
            self._batch_btn.setStyleSheet(
                f"QPushButton {{ background: {t['danger']}; color: white; border: none;"
                f" border-radius: 8px; font-size: 13px; font-weight: 600; }}"
                f"QPushButton:hover {{ background: {t['danger']}; }}"
            )
        else:
            self._batch_btn.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {t['text_body']};"
                f" border: 1px solid {t['border']}; border-radius: 8px; font-size: 13px; font-weight: 600; }}"
                f"QPushButton:hover {{ border-color: {t['accent']}; color: {t['text_hi']}; }}"
            )

    def apply_theme(self):
        """主题切换时整页同步：静态控件重上色 + 列表卡片重渲染。
        卡片样式是创建时的快照，不重渲染会深浅色混杂。"""
        self._style_search_edit()
        self._sync_filter_style()
        if hasattr(self, "_batch_btn"):
            self._style_batch_btn()
        if self._entries:
            self._render()

    def _clear_list(self):
        """清空列表：先隐藏再销毁，避免旧卡片在删除延迟期间与新卡片同时可见"""
        while self.list_layout.count() > 1:
            item = self.list_layout.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

    def _toggle_batch_mode(self):
        self._batch_mode = self._batch_btn.isChecked()
        self._selected_paths.clear()
        self._batch_btn.setText("✅ 退出批量" if self._batch_mode else "🗑️ 批量删除")
        self._style_batch_btn()
        self._page = 1
        self._render()

    def _on_search_changed(self, text):
        self._search_text = (text or "").strip().lower()
        self._page = 1
        self._render()

    def _set_filter(self, key):
        self._filter = key
        self._page = 1
        self._sync_filter_style()
        self._render()

    def _sync_filter_style(self):
        t = get_tokens(self._is_dark())
        if hasattr(self, "search_edit"):
            self._style_search_edit()
        for key, btn in self._filter_buttons.items():
            checked = key == self._filter
            btn.setChecked(checked)
            btn.setStyleSheet(
                f"QPushButton {{ border: 1px solid {'#3b82f6' if checked else t['border']}; background: {'#3b82f6' if checked else 'transparent'}; color: {'white' if checked else t['text_body']}; border-radius: 8px; font-size: 13px; font-weight: 600; }} QPushButton:hover {{ border-color: {t['accent']}; color: {t['text_hi']}; }}")
        if hasattr(self, "_refresh_btn"):
            self._refresh_btn.setStyleSheet(
                f"QPushButton {{ border: 1px solid {t['border']}; background: transparent; color: {t['text_body']}; border-radius: 8px; font-size: 13px; font-weight: 600; }} QPushButton:hover {{ border-color: {t['accent']}; color: {t['text_hi']}; }}")

    def refresh(self):
        """刷新历史列表（异步化）"""
        # 【核心保护】大型项目运行期间拦截磁盘扫描，避免 IO 争抢导致界面卡顿；
        # 任务结束后由定时器自动恢复刷新
        if is_heavy_task_running():
            self._show_heavy_task_warning()
            return
        self._clear_heavy_task_warning()

        self._merge_enabled = MatcherConfig.load().get("history_display_mode", "merge") == "merge"

        self._refresh_btn.setEnabled(False)
        self._refresh_btn.setText("⏳ 加载中...")

        self._clear_list()

        loading_lbl = QLabel("⏳ 正在扫描历史项目，请稍候...")
        loading_lbl.setAlignment(Qt.AlignCenter)
        loading_lbl.setStyleSheet(f"color: {self._sub_color()}; font-size: 14px; padding: 60px 0;")
        self.list_layout.insertWidget(self.list_layout.count() - 1, loading_lbl)

        # 启动后台扫描线程
        # 【崩溃修复】旧扫描线程可能仍在运行，直接覆盖引用会销毁运行中的
        # QThread → Qt qFatal 闪退。改为移入退休列表暂存，结束后再清理。
        old_worker = getattr(self, "_scan_worker", None)
        if old_worker is not None:
            if not hasattr(self, "_retired_scan_workers"):
                self._retired_scan_workers = []
            self._retired_scan_workers.append(old_worker)
            # 只保留仍在运行的退休线程；清理陈旧信号连接避免旧结果覆盖新列表
            self._retired_scan_workers = [
                w for w in self._retired_scan_workers if w.isRunning()
            ]
            try:
                old_worker.finished.disconnect(self._on_scan_finished)
            except Exception:
                pass
        self._scan_worker = HistoryScanWorker()
        self._scan_worker.finished.connect(self._on_scan_finished)
        self._scan_worker.start()

    def _on_scan_finished(self, entries):
        """后台扫描完成回调（在主线程执行）"""
        self._entries = entries
        self._page = 1

        self._refresh_btn.setEnabled(True)
        self._refresh_btn.setText("🔄 刷新")

        self._render()

    def _show_heavy_task_warning(self):
        """大型项目处理中：显示降级提示，并定时检测任务结束自动恢复"""
        self._refresh_btn.setEnabled(True)
        self._refresh_btn.setText("🔄 刷新")

        # 清空列表，仅保留末尾的固定控件
        self._clear_list()

        warning_lbl = QLabel(
            "⚠️ 系统正在集中资源处理大型项目（>50MB），历史列表加载已暂缓。\n"
            "任务完成后将自动恢复，也可稍后手动点击刷新。"
        )
        warning_lbl.setObjectName("HeavyTaskWarning")
        warning_lbl.setAlignment(Qt.AlignCenter)
        warning_lbl.setStyleSheet(
            "color: #f59e0b; font-size: 14px; font-weight: bold; padding: 40px 20px; "
            "background: rgba(245, 158, 11, 0.1); border-radius: 8px; "
            "border: 1px dashed #f59e0b;"
        )
        self.list_layout.insertWidget(self.list_layout.count() - 1, warning_lbl)

        # 每 10 秒检测一次大任务是否结束
        if getattr(self, "_heavy_task_timer", None) is None:
            self._heavy_task_timer = QTimer(self)
            self._heavy_task_timer.setInterval(10000)
            self._heavy_task_timer.timeout.connect(self._check_and_auto_refresh)
        self._heavy_task_timer.start()

    def _clear_heavy_task_warning(self):
        """移除大任务提示控件并停止自动恢复定时器"""
        timer = getattr(self, "_heavy_task_timer", None)
        if timer is not None and timer.isActive():
            timer.stop()
        for i in range(self.list_layout.count()):
            item = self.list_layout.itemAt(i)
            w = item.widget() if item else None
            if w is not None and w.objectName() == "HeavyTaskWarning":
                self.list_layout.removeWidget(w)
                w.deleteLater()
                break

    def _check_and_auto_refresh(self):
        """定时检测大任务是否结束，结束后自动恢复刷新"""
        if not is_heavy_task_running():
            self.refresh()

    def _get_date_group(self, ts):
        """改进的日期分组逻辑，支持本月/上月"""
        try:
            dt = datetime.strptime(ts, "%Y%m%d_%H%M%S")
            now = datetime.now()
            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
            delta_days = (today - dt.replace(hour=0, minute=0, second=0, microsecond=0)).days

            if delta_days == 0: return "今天"
            if delta_days == 1: return "昨天"
            if delta_days <= 6: return "本周"
            if delta_days <= 13: return "上周"

            if dt.year == now.year and dt.month == now.month:
                return "本月"
            elif (dt.year == now.year and dt.month == now.month - 1) or (
                    now.month == 1 and dt.year == now.year - 1 and dt.month == 12):
                return "上月"
            return "更早"
        except:
            return "更早"

    def _render(self):
        self._clear_list()

        visible = [e for e in self._entries
                   if self._filter in ("全部", e["type"])
                   and (not self._search_text or self._search_text in str(e["project"]).lower())]

        if self._merge_enabled:
            groups = defaultdict(list)
            for e in visible: groups[e["project"]].append(e)
            sorted_projects = sorted(groups.keys(), key=lambda p: groups[p][0]["ts"], reverse=True)
            total_pages = max(1, (len(sorted_projects) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        else:
            sorted_projects = None

            # 【核心修复】：非合并模式下，先按日期分组排序，再计算分页
            def sort_key(e):
                group = self._get_date_group(e["ts"])
                group_idx = _GROUP_ORDER.index(group) if group in _GROUP_ORDER else 99
                # 组索引升序（今天->更早），时间戳降序（最新->最旧）
                # 将时间戳转为整数并取负，实现组内倒序
                ts_num = int(e["ts"].replace("_", "")) if e["ts"] else 0
                return (group_idx, -ts_num)

            visible = sorted(visible, key=sort_key)
            total_pages = max(1, (len(visible) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)

        self._page = min(max(1, self._page), total_pages)
        sub_color = self._sub_color()

        count_label = QLabel(f"共 {len(visible)} 条记录 · 第 {self._page}/{total_pages} 页")
        count_label.setStyleSheet(f"color: {sub_color}; font-size: 12px; padding-left: 4px;")
        self.list_layout.insertWidget(self.list_layout.count() - 1, count_label)

        if not visible:
            empty = QLabel("暂无历史记录")
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(f"color: {sub_color}; font-size: 14px; padding: 60px 0;")
            self.list_layout.insertWidget(self.list_layout.count() - 1, empty)
            return

        start = (self._page - 1) * self.PAGE_SIZE

        if self._merge_enabled and sorted_projects:
            for project_name in sorted_projects[start: start + self.PAGE_SIZE]:
                self.list_layout.insertWidget(
                    self.list_layout.count() - 1,
                    HistoryGroupCard(project_name, groups[project_name], parent=self, batch_mode=self._batch_mode)
                )
        else:
            # 对排序后的列表进行切片
            page_entries = visible[start: start + self.PAGE_SIZE]
            current_group = None
            for e in page_entries:
                group_name = self._get_date_group(e["ts"])
                if group_name != current_group:
                    current_group = group_name
                    group_header = QLabel(f"📅 {group_name}")
                    group_header.setStyleSheet(
                        f"font-size: 14px; font-weight: 700; color: {sub_color}; padding: 15px 0 8px 0;")
                    self.list_layout.insertWidget(self.list_layout.count() - 1, group_header)
                self.list_layout.insertWidget(self.list_layout.count() - 1, self._make_flat_card(e))

        self._add_pagination(total_pages, sub_color)

    def _add_pagination(self, total_pages, sub_color):
        nav = QHBoxLayout()
        nav.setSpacing(10)

        t = get_tokens(self._is_dark())
        _nav_qss = (f"QPushButton {{ border: 1px solid {t['border']}; background: transparent;"
                    f" color: {t['text_body']}; border-radius: 6px; padding: 0 14px; font-size: 13px; }}"
                    f"QPushButton:hover {{ border-color: {t['accent']}; color: {t['text_hi']}; }}"
                    "QPushButton:disabled { color: rgba(128,128,128,0.5); }")

        # 【新增】第一页按钮
        btn_first = QPushButton("⏮ 第一页")
        btn_first.setFixedHeight(30)
        btn_first.setCursor(Qt.PointingHandCursor)
        btn_first.setStyleSheet(_nav_qss)
        btn_first.setEnabled(self._page > 1)
        btn_first.clicked.connect(lambda: self._goto_page(1))
        nav.addWidget(btn_first)

        btn_prev = QPushButton("← 上一页")
        btn_prev.setFixedHeight(30)
        btn_prev.setCursor(Qt.PointingHandCursor)
        btn_prev.setStyleSheet(_nav_qss)
        btn_prev.setEnabled(self._page > 1)
        btn_prev.clicked.connect(lambda: self._goto_page(self._page - 1))
        nav.addWidget(btn_prev)

        page_label = QLabel(f"{self._page} / {total_pages}")
        page_label.setStyleSheet(f"color: {sub_color}; font-size: 13px; font-weight: 600;")
        page_label.setAlignment(Qt.AlignCenter)
        nav.addWidget(page_label)

        btn_next = QPushButton("下一页 →")
        btn_next.setFixedHeight(30)
        btn_next.setCursor(Qt.PointingHandCursor)
        btn_next.setStyleSheet(_nav_qss)
        btn_next.setEnabled(self._page < total_pages)
        btn_next.clicked.connect(lambda: self._goto_page(self._page + 1))
        nav.addWidget(btn_next)

        # 【新增】最后一页按钮
        btn_last = QPushButton("最后一页 ⏭")
        btn_last.setFixedHeight(30)
        btn_last.setCursor(Qt.PointingHandCursor)
        btn_last.setStyleSheet(_nav_qss)
        btn_last.setEnabled(self._page < total_pages)
        btn_last.clicked.connect(lambda: self._goto_page(total_pages))
        nav.addWidget(btn_last)

        nav.addStretch()
        nav_widget = QWidget()
        nav_widget.setLayout(nav)
        nav_widget.setStyleSheet("background: transparent;")
        self.list_layout.insertWidget(self.list_layout.count() - 1, nav_widget)

    def _goto_page(self, page):
        self._page = max(1, page)
        self._render()

    def _make_flat_card(self, e):
        meta = _TYPE_META[e["type"]]
        card = QFrame()
        card.setObjectName("HistoryCard")
        is_dark = self._is_dark()
        t = get_tokens(is_dark)
        card.setStyleSheet(
            f"QFrame#HistoryCard {{ background: {t['bg_card']}; border: 1px solid {t['border']}; border-left: 4px solid {meta['color']}; border-radius: 10px; }} QFrame#HistoryCard:hover {{ background: {t['bg_hover']}; }}")

        main_layout = QVBoxLayout(card)
        main_layout.setContentsMargins(16, 12, 16, 12)
        main_layout.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.setSpacing(12)

        if self._batch_mode:
            cb = QCheckBox()
            cb.setCursor(Qt.PointingHandCursor)
            is_selected = e.get("path") in self._selected_paths
            cb.setChecked(is_selected)
            cb.setStyleSheet(
                "QCheckBox { spacing: 5px; } QCheckBox::indicator { width: 18px; height: 18px; border: 2px solid #94a3b8; border-radius: 4px; background-color: transparent; } QCheckBox::indicator:checked { background-color: #3b82f6; border: 2px solid #3b82f6; } QCheckBox::indicator:hover { border: 2px solid #3b82f6; }")
            cb.toggled.connect(lambda checked, p=e.get("path"): self._on_item_selected(p, checked))
            top_row.addWidget(cb)

        icon = QLabel(meta["icon"])
        icon.setStyleSheet("font-size: 22px;")
        top_row.addWidget(icon)

        badge = QLabel(e["type"])
        badge.setStyleSheet(
            f"background: {meta['color']}; color: white; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: 700;")
        badge.setFixedHeight(22)
        top_row.addWidget(badge)

        name = ElidedTitleLabel(str(e["project"]))
        name.setStyleSheet("font-size: 14px; font-weight: 700;")
        top_row.addWidget(name, 1)

        status_text, status_color, stats = e.get("status_info", ("完成", "#10b981", None))
        if e.get("paused"):
            paused_data = _load_json_safe(e["paused"])
            if paused_data:
                status_text, status_color, stats = "暂停", "#f59e0b", f"进度 {paused_data.get('completed_through', 0)}/10"

        status_container = QWidget()
        status_layout = QHBoxLayout(status_container)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(8)
        status_layout.setAlignment(Qt.AlignRight)

        if stats:
            stats_label = QLabel(stats)
            stats_label.setStyleSheet(
                f"color: {self._sub_color()}; font-size: 12px; font-weight: 500; padding: 2px 6px;")
            status_layout.addWidget(stats_label)

        status_label = QLabel(status_text)
        status_label.setStyleSheet(
            f"color: {status_color}; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 12px; background: {self._get_soft_bg_color(status_color)}; border: 1px solid {self._get_border_color(status_color)};")
        status_layout.addWidget(status_label)
        top_row.addWidget(status_container)
        main_layout.addLayout(top_row)

        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(10)
        info_label = QLabel(f"{e['time_text']}  ·  {e['detail']}")
        info_label.setStyleSheet(f"color: {self._sub_color()}; font-size: 12px;")
        bottom_row.addWidget(info_label)
        bottom_row.addStretch()

        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        if e.get("type") == "初评" and e.get("run_info"):
            btn_rerun = QPushButton("🔁 重新初评")
            btn_rerun.setFixedHeight(30);
            btn_rerun.setFixedWidth(90);
            btn_rerun.setCursor(Qt.PointingHandCursor)
            btn_rerun.setStyleSheet(
                f"font-size: 12px; padding: 0 8px; border: 1px solid {t['border']}; border-radius: 6px; background: transparent; color: {t['text_body']};")
            btn_rerun.clicked.connect(lambda _=False, p=e.get("run_info"): self._emit_re_run(p))
            btn_layout.addWidget(btn_rerun)

        if e.get("paused"):
            btn_resume = QPushButton("▶️ 继续")
            btn_resume.setFixedHeight(30);
            btn_resume.setFixedWidth(70);
            btn_resume.setCursor(Qt.PointingHandCursor)
            btn_resume.setStyleSheet(
                "font-size: 12px; padding: 0 8px; background: #10b981; color: white; border: none; border-radius: 6px; font-weight: 600;")
            btn_resume.clicked.connect(lambda _=False, p=e.get("paused"): self._emit_resume(p))
            btn_layout.addWidget(btn_resume)

        btn_open = QPushButton("📂 打开文件夹")
        btn_open.setFixedHeight(30);
        btn_open.setFixedWidth(100);
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.setStyleSheet(
            f"font-size: 12px; padding: 0 8px; border: 1px solid {t['border']}; border-radius: 6px; background: transparent; color: {t['text_body']};")
        btn_open.clicked.connect(lambda _=False, p=e.get("path"): self._open_dir(p))
        btn_layout.addWidget(btn_open)

        bottom_row.addWidget(btn_container)
        card.setContextMenuPolicy(Qt.CustomContextMenu)
        card.customContextMenuRequested.connect(lambda pos, entry=e: self._show_card_context_menu(pos, entry, card))
        main_layout.addLayout(bottom_row)
        return card

    def _on_item_selected(self, path, checked):
        if checked:
            self._selected_paths.add(path)
        else:
            self._selected_paths.discard(path)

    def _get_soft_bg_color(self, status_color):
        if status_color == "#ef4444":
            return "rgba(239, 68, 68, 0.08)"
        elif status_color == "#10b981":
            return "rgba(16, 185, 129, 0.08)"
        elif status_color == "#f59e0b":
            return "rgba(245, 158, 11, 0.08)"
        return "rgba(148, 163, 184, 0.08)"

    def _get_border_color(self, status_color):
        if status_color == "#ef4444":
            return "rgba(239, 68, 68, 0.2)"
        elif status_color == "#10b981":
            return "rgba(16, 185, 129, 0.2)"
        elif status_color == "#f59e0b":
            return "rgba(245, 158, 11, 0.2)"
        return "rgba(148, 163, 184, 0.2)"

    def _show_card_context_menu(self, pos, entry, widget):
        menu = QMenu(self)
        # 菜单样式跟随主题（原先硬编码深色，浅色模式下黑底突兀）
        t = get_tokens(self._is_dark())
        menu.setStyleSheet(
            f"QMenu {{ background-color: {t['bg_card']}; color: {t['text_body']};"
            f" border: 1px solid {t['border']}; border-radius: 8px; padding: 4px; }}"
            f"QMenu::item {{ padding: 6px 24px 6px 12px; border-radius: 4px; }}"
            f"QMenu::item:selected {{ background-color: {t['accent']}; color: #ffffff; }}"
            f"QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 8px; }}")

        act_open_dir = QAction("📂 打开产物文件夹", self)
        act_open_dir.triggered.connect(lambda: self._open_dir(entry.get("path")))
        menu.addAction(act_open_dir)

        if entry.get("upload_manifest"):
            act_open_upload = QAction(" 打开上传文件位置", self)
            act_open_upload.triggered.connect(lambda: self._open_upload_location(entry))
            menu.addAction(act_open_upload)

        if entry.get("log_file"):
            act_open_log = QAction("📄 打开运行日志", self)
            act_open_log.triggered.connect(lambda: self._open_file(entry.get("log_file")))
            menu.addAction(act_open_log)

        if entry.get("type") == "初评":
            act_show_summary = QAction("📊 查看结果汇总", self)
            act_show_summary.triggered.connect(lambda: self._show_saved_summary(entry))
            menu.addAction(act_show_summary)

        menu.addSeparator()
        if entry.get("type") == "初评":
            act_rerun = QAction("🔁 重新评估", self)
            act_rerun.triggered.connect(lambda: self._emit_re_run(entry.get("run_info")))
            menu.addAction(act_rerun)
        if entry.get("paused"):
            act_resume = QAction("▶️ 继续评估", self)
            act_resume.triggered.connect(lambda: self._emit_resume(entry.get("paused")))
            menu.addAction(act_resume)
        menu.addSeparator()

        if self._batch_mode:
            count = len(self._selected_paths)
            act_batch_delete = QAction(f"🗑️ 删除选中的 {count} 项", self)
            act_batch_delete.setEnabled(count > 0)
            act_batch_delete.triggered.connect(self._delete_selected_entries)
            menu.addAction(act_batch_delete)
        else:
            act_delete = QAction("🗑️ 删除记录", self)
            act_delete.triggered.connect(lambda: self._delete_single_entry(entry))
            menu.addAction(act_delete)

        menu.exec(widget.mapToGlobal(pos))

    def _open_file(self, path):
        if path and os.path.exists(path):
            try:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", path])
            except Exception as e:
                log_error(f'打开文件失败: {e}')

    def _open_upload_location(self, entry):
        import subprocess
        target = self._resolve_upload_file(entry)
        if not target:
            self._open_dir(entry.get("path", ""))
            return
        try:
            if os.name == "nt":
                subprocess.Popen(["explorer", "/select,", os.path.normpath(target)])
            else:
                self._open_dir(os.path.dirname(target))
        except Exception as exc:
            log_error(f'打开材料位置失败: {exc}')

    def _resolve_upload_file(self, entry):
        manifest = entry.get("upload_manifest")
        if not manifest or not os.path.exists(manifest): return None
        data = self._load_json(manifest)
        if not data: return None
        candidates = []
        etype = entry.get("type")
        if etype == "初评":
            for pair in data.get("file_pairs") or []:
                if isinstance(pair, dict): candidates += [pair.get("word"), pair.get("excel"), pair.get("asset")]
        elif etype == "重评":
            candidates = [data.get("excel2"), data.get("excel1")]
        else:
            candidates = [data.get("eval_report_path"), data.get("eval_consent_path")]
            for tk in data.get("tasks") or []:
                if isinstance(tk, dict): candidates += [tk.get("eval_report"), tk.get("eval_consent")]
        for cand in candidates:
            if cand and os.path.exists(str(cand)): return str(cand)
        return None

    def _load_json(self, path):
        import json
        if not path or not os.path.exists(path): return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return None

    def _delete_single_entry(self, entry):
        reply = QMessageBox.question(self, "确认删除",
                                     f"确定删除该记录及其文件吗？\n项目：{entry.get('project')}\n时间：{entry.get('time_text')}",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            path = entry.get("path")
            if path and os.path.exists(path):
                try:
                    shutil.rmtree(path)
                    log_info(f'[DELETE] ✅ 已删除时间戳目录：{path}')
                    self._cleanup_project_if_last(entry)
                    self.refresh()
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"删除失败：{e}")

    def _delete_selected_entries(self):
        if not self._selected_paths: return
        reply = QMessageBox.question(self, "确认批量删除",
                                     f"确定删除选中的 {len(self._selected_paths)} 条记录及其文件吗？\n此操作不可恢复！",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            deleted_entries = []
            for path in list(self._selected_paths):
                if path and os.path.exists(path):
                    try:
                        shutil.rmtree(path)
                        log_info(f'[DELETE] ✅ 已删除：{path}')
                        for e in self._entries:
                            if e.get("path") == path:
                                deleted_entries.append(e)
                                break
                    except Exception as e:
                        log_error(f'[DELETE] 删除 {path} 失败: {e}')
            for entry in deleted_entries:
                self._cleanup_project_if_last(entry)
            self._selected_paths.clear()
            self.refresh()

    def _emit_re_run(self, path):
        data = _load_json_safe(path) if path else None
        if path and not data: QMessageBox.warning(self, "历史配置已损坏",
                                                  "该历史任务的配置文件已损坏，将引导手动选择材料。")
        if not data or not data.get("file_pairs"): data = self._manual_pick_materials()
        if not data: return
        data.pop("resume_state", None);
        data.pop("resume_source_paused_path", None)
        self.re_run_requested.emit(data)

    def _manual_pick_materials(self):
        from PySide6.QtWidgets import QFileDialog
        word, _ = QFileDialog.getOpenFileName(self, "选择需求说明书 Word", "", "Word 文件 (*.docx *.doc)")
        if not word: return None
        excel, _ = QFileDialog.getOpenFileName(self, "选择拆分表 Excel", "", "Excel 文件 (*.xlsx *.xls)")
        if not excel: return None
        from utils.path_utils import clean_project_name
        name = clean_project_name(os.path.splitext(os.path.basename(word))[0])
        return {"filename": name, "days": "0", "file_pairs": [{"word": word, "excel": excel, "asset": None}],
                "auto_numbering": False, "run_template": True, "run_empty": True, "run_ratio": True,
                "run_factors": True, "run_hierarchy": True, "run_simple": False, "run_move": True, "run_asset": False,
                "run_data_attribute_check": False, "validation_results": []}

    def _emit_resume(self, path):
        data = _load_json_safe(path)
        if data is None: return
        if path and os.path.exists(path):
            try:
                os.remove(path)
                log_info(f'[CLEANUP] 已清除旧的暂停状态文件: {path}')
            except Exception as e:
                log_warn(f'[WARN] 清除暂停状态文件失败: {e}')

        raw = dict(data.get("raw_task_info", {}) or {})
        raw["resume_state"] = {"v_res": data.get("v_res", {}), "completed_through": data.get("completed_through", -1)}
        raw["resume_source_paused_path"] = path
        self.resume_requested.emit(raw)

    @staticmethod
    def _open_dir(path):
        try:
            if path and os.path.isdir(path):
                if os.name == "nt":
                    os.startfile(path)
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            log_error(f'打开文件夹失败: {exc}')

    def _show_saved_summary(self, entry):
        from ui.report_dialog import SummaryDialog

        summary_path = entry.get("summary_html_path") or entry.get("summary_html")
        html_content = ""
        if isinstance(summary_path, str) and os.path.exists(summary_path):
            try:
                with open(summary_path, "r", encoding="utf-8") as f:
                    html_content = f.read()
            except Exception as e:
                log_error(f'读取 HTML 文件失败: {e}')
        elif isinstance(summary_path, str):
            html_content = summary_path

        if html_content:
            dlg = SummaryDialog(task_data={}, summary_html=html_content, parent=self)
            dlg.exec()
            return

        # 兜底：无存档 HTML 时，从 result.json 的 v_res 现场重新生成汇总
        result_path = os.path.join(entry.get("path", ""), "result.json")
        res_data = self._load_json(result_path) or {}
        v_res = res_data.get("v_res") or {}
        if not v_res:
            QMessageBox.warning(self, "提示", "未找到汇总报告内容。")
            return
        task_data = {"filename": entry.get("project", ""), "validation_results": [v_res]}
        dlg = SummaryDialog(task_data=task_data, parent=self)
        dlg.exec()

    def _cleanup_project_if_last(self, entry):
        path = entry.get("path")
        if not path: return
        project_dir = os.path.dirname(path)
        if not os.path.exists(project_dir): return

        remaining_ts_dirs = []
        try:
            for item in os.listdir(project_dir):
                item_path = os.path.join(project_dir, item)
                if os.path.isdir(item_path) and _parse_ts_from_name(item):
                    remaining_ts_dirs.append(item)
        except Exception as e:
            log_error(f'[CLEANUP] 检查项目文件夹内容失败: {e}')

        if not remaining_ts_dirs:
            log_info(f'[CLEANUP] 最后一个记录，开始清理项目文件夹和缓存: {project_dir}')
            try:
                for f in os.listdir(project_dir):
                    if f.endswith("-tree_cache.json"):
                        os.remove(os.path.join(project_dir, f))
            except Exception as e:
                log_error(f'[CLEANUP] 删除缓存文件失败: {e}')

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    shutil.rmtree(project_dir)
                    if not os.path.exists(project_dir):
                        log_info(f'[CLEANUP] ✅ 已删除项目文件夹: {project_dir}')
                        break
                except PermissionError as e:
                    if attempt < max_retries - 1:
                        import time;
                        time.sleep(1)
                except Exception as e:
                    if attempt < max_retries - 1:
                        import time;
                        time.sleep(1)