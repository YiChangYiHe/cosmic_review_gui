#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx 结构树快速提取器（XML 直读单遍扫描）

背景：超大 docx（300MB+，几十万段落）走 python-docx 物理遍历或 COM 逐段读取
（每段 3+ 次跨进程往返）需要 1~2 小时。本模块直接用 iterparse 流式解析
word/document.xml，一次遍历同时产出：
  1. 标题大纲 outline: [(number, title, level), ...]（编号推导逻辑与
     HierarchicalMatcher.extract_word_outline_as_hierarchy 完全一致，
     保证结果兼容）
  2. 全文流 full_flow（与 HierarchicalMatcher._read_word_full_text 输出同构，
     供功能过程章节匹配使用）

仅依赖标准库 zipfile + lxml（xml.etree 兜底），PyInstaller 打包无新增依赖。
.doc 老格式由调用方先转换为 .docx（复用 convert_doc_to_docx）。
"""

import re
import zipfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:  # 打包环境已含 lxml；此处仅作极端兜底
    from xml.etree import ElementTree as etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

# 与 extract_word_outline_as_hierarchy 保持一致的手工编号匹配
_MANUAL_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)[\s\.．、)）:：]*")
_STYLE_LEVEL_RE = re.compile(r"(?:heading|标题)\s*(\d)", re.IGNORECASE)
_TOC_DOTS_RE_TAIL = re.compile(r"\t\s*\d+\s*$")
_TOC_DOTS_RE = re.compile(r"\.{2,}\s*\d+\s*$")
_BODY_SKIP_KEYWORDS = ("请在此处", "说明本项目", "示例（", "示例:")


def _is_body_paragraph(title: str) -> bool:
    """过滤误标大纲级别的正文段落（与主提取器一致）"""
    t = (title or "").strip()
    if not t:
        return True
    if re.match(r"^[①②③④⑥⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]", t):
        return True
    if re.match(r"^[(（]\s*\d+\s*[)）]", t):
        return True
    if len(t) > 60:
        return True
    if t.endswith(("。", "；", "，", "！", "？", ";", ",")):
        return True
    return False


def _parse_styles(zf):
    """解析 word/styles.xml：styleId -> (name, outline_level)"""
    style_map = {}
    try:
        data = zf.read("word/styles.xml")
    except KeyError:
        return style_map
    try:
        root = etree.fromstring(data)
    except Exception:
        return style_map
    for st in root.iter(W + "style"):
        if st.get(W + "type") != "paragraph":
            continue
        sid = st.get(W + "styleId") or ""
        if not sid:
            continue
        name = ""
        name_el = st.find(W + "name")
        if name_el is not None:
            name = name_el.get(W + "val") or ""
        lvl = None
        ppr = st.find(W + "pPr")
        if ppr is not None:
            ol = ppr.find(W + "outlineLvl")
            if ol is not None:
                try:
                    lvl = int(ol.get(W + "val")) + 1
                except (TypeError, ValueError):
                    lvl = None
        if lvl is None and name:
            m = _STYLE_LEVEL_RE.search(name)
            if m:
                lvl = int(m.group(1))
        style_map[sid] = (name, lvl)
    return style_map


def _iter_docx_paragraphs(docx_path, progress_cb=None, progress_every=20000):
    """
    流式遍历 document.xml（含表格内段落），按文档顺序产出段落信息字典：
      {"text", "style_id", "outline_level", "in_table"}
    逐元素清理内存，300MB+ 文档内存占用可控。
    """
    with zipfile.ZipFile(docx_path) as zf:
        try:
            stream = zf.open("word/document.xml")
        except KeyError:
            return
        para_count = 0
        in_table = 0
        cur = None
        for event, el in etree.iterparse(stream, events=("start", "end")):
            tag = el.tag
            if event == "start":
                if tag == W + "tbl":
                    in_table += 1
                elif tag == W + "p":
                    cur = {
                        "text_parts": [],
                        "style_id": "",
                        "outline_level": None,
                        "in_table": in_table > 0,
                    }
            else:  # end
                if tag == W + "tbl":
                    in_table -= 1
                elif cur is not None:
                    if tag == W + "t":
                        if el.text:
                            cur["text_parts"].append(el.text)
                    elif tag == W + "tab":
                        cur["text_parts"].append("\t")
                    elif tag in (W + "br", W + "cr"):
                        cur["text_parts"].append("\n")
                    elif tag == W + "pStyle":
                        v = el.get(W + "val")
                        if v and not cur["style_id"]:
                            cur["style_id"] = v
                    elif tag == W + "outlineLvl":
                        try:
                            cur["outline_level"] = int(el.get(W + "val")) + 1
                        except (TypeError, ValueError):
                            pass
                    elif tag == W + "p":
                        para_count += 1
                        if progress_cb and para_count % progress_every == 0:
                            progress_cb(para_count)
                        info = {
                            "text": "".join(cur["text_parts"]),
                            "style_id": cur["style_id"],
                            "outline_level": cur["outline_level"],
                            "in_table": cur["in_table"],
                        }
                        cur = None
                        # 释放已处理元素，防止超大文档内存膨胀
                        el.clear()
                        parent = el.getparent()
                        if parent is not None:
                            while el.getprevious() is not None:
                                del parent[0]
                        yield info


def _flow_item(text, cleaner, source, parent_num="", number=""):
    return {
        "text": text,
        "original": text,
        "cleaned": cleaner.basic_clean(text),
        "source": source,
        "type": source,
        "parent_num": parent_num,
        "number": number,
        "in_toc": source == "heading",
    }


def _build_headings_queue(word_items, cleaner):
    """由标题项构建顺序对齐队列（与 _read_word_full_text 一致）"""
    headings_queue = []
    for item in word_items or []:
        num = item.get("number", "")
        raw = (item.get("original") or item.get("text", "")).strip().lower()
        clean = cleaner.basic_clean(raw).lower()
        title_only = (item.get("title") or item.get("display") or "").strip()
        title_clean = re.sub(
            r"^\d+(?:[\.．]\d+)*[\.．、)）:：\s\u3000-]*",
            "",
            cleaner.basic_clean(title_only),
        ).lower()
        if num:
            headings_queue.append(
                {
                    "number": num,
                    "raw": raw,
                    "clean": clean,
                    "title_clean": title_clean,
                }
            )
    return headings_queue


def _flow_pass(docx_path, headings_queue, cleaner, progress_cb=None):
    """
    单遍全文流提取。headings_queue 非空时按文本对齐标题（与
    _read_word_full_text 行为一致）；为空时按顺序消耗队列之外的标题。
    """
    flow = []
    current_num = ""
    heading_ptr = 0
    num_headings = len(headings_queue)

    for info in _iter_docx_paragraphs(docx_path, progress_cb=progress_cb):
        text = info["text"].strip()
        if not text:
            continue

        style_name = ""
        # 表格文本（不参与标题判定）
        if info["in_table"]:
            if len(text) > 1:
                flow.append(_flow_item(text, cleaner, "table", parent_num=current_num))
            continue

        # 样式名（用于目录过滤）
        # style 名称需调用方传入太重，这里仅依赖 style_id 由调用方解析；
        # 目录段落通过 TOC 样式 id 或尾随页码特征过滤
        sid = info["style_id"]
        if sid and (sid.lower().startswith("toc") or "目录" in sid):
            continue
        if "toc" in text[:12].lower() and _TOC_DOTS_RE.search(text):
            continue
        if _TOC_DOTS_RE_TAIL.search(text) or _TOC_DOTS_RE.search(text):
            continue

        if headings_queue and heading_ptr < num_headings:
            p_lower = text.lower()
            p_clean = cleaner.basic_clean(text).lower()
            p_no_num = re.sub(
                r"^\d+(?:[\.．]\d+)*[\.．、)）:：\s\u3000-]*", "", p_clean
            )
            lookahead_limit = 1 if len(p_clean) < 10 else 5
            matched = False
            for offset in range(min(lookahead_limit, num_headings - heading_ptr)):
                target = headings_queue[heading_ptr + offset]
                if (
                    p_lower == target["raw"]
                    or p_clean == target["clean"]
                    or p_lower == target["clean"]
                    or (p_no_num and p_no_num == target.get("title_clean"))
                ):
                    current_num = target["number"]
                    heading_ptr += offset + 1
                    matched = True
                    break
            if matched:
                flow.append(_flow_item(text, cleaner, "heading", number=current_num))
                continue
        elif info["outline_level"] and 1 <= info["outline_level"] <= 9:
            # 无队列时按大纲级别直接认定标题
            current_num = ""
            flow.append(_flow_item(text, cleaner, "heading", number=""))
            continue

        if len(text) > 2 and not any(kw in text for kw in _BODY_SKIP_KEYWORDS):
            flow.append(_flow_item(text, cleaner, "paragraph", parent_num=current_num))

    return flow


def _strip_leading_zero_parts(num: str):
    """
    去掉编号前缀中的 0 段，还原实际编号：
      "0.4.1" -> "4.1"   "0.4.1.1" -> "4.1.1"   "0.1" -> "1"
    0 段来自顶级标题未被识别时的层级计数器占位；全部为 0 时保留原值。
    返回 (新编号, 是否发生了剥离)。
    """
    parts = num.split(".")
    i = 0
    while i < len(parts) - 1 and parts[i] == "0":
        i += 1
    if i == 0:
        return num, False
    return ".".join(parts[i:]), True


def extract_doc_tree(
    docx_path,
    cleaner,
    auto_numbering=False,
    max_level=9,
    include_full_text=True,
    progress_cb=None,
):
    """
    单遍提取标题大纲 + 可选全文流。
    返回 (outline, full_flow)；outline 为 [(number, title, level), ...]。
    编号推导逻辑与 extract_word_outline_as_hierarchy 完全一致：
      - auto_numbering=False：优先取标题文字中的手工编号，缺失时用层级计数器
      - auto_numbering=True：全部使用层级计数器（与现有行为一致）
    解析失败抛异常，由调用方回退旧通道。
    """
    path = str(docx_path)
    with zipfile.ZipFile(path) as zf:
        style_map = _parse_styles(zf)

    outline = []
    full_flow = [] if include_full_text else None
    current_num = ""
    auto_counters = [0] * 10

    for info in _iter_docx_paragraphs(docx_path, progress_cb=progress_cb):
        text = info["text"]
        if not isinstance(text, str):
            # 防御：异常文档中 text 可能被污染为非字符串
            text = "" if text is None else str(text)
        if not text or not text.strip():
            continue

        if info["in_table"]:
            if include_full_text and len(text.strip()) > 1:
                full_flow.append(
                    _flow_item(text.strip(), cleaner, "table", parent_num=current_num)
                )
            continue

        stripped = text.strip()

        # 目录段落过滤（样式 id：TOC1/目录 1 等；以及点线+页码特征）
        sid = info["style_id"]
        if sid and (sid.lower().startswith("toc") or "目录" in sid):
            continue
        if _TOC_DOTS_RE_TAIL.search(text) or _TOC_DOTS_RE.search(text):
            continue

        # 大纲级别：段落直接 outlineLvl > 样式表 outlineLvl > 样式名正则
        outline_level = info["outline_level"]
        style_name, style_lvl = style_map.get(info["style_id"], ("", None))
        if outline_level is None:
            outline_level = style_lvl
        if outline_level is None and style_name:
            m = _STYLE_LEVEL_RE.search(style_name)
            if m:
                outline_level = int(m.group(1))

        # ---------- 全文流（与 _read_word_full_text 行为一致） ----------
        if include_full_text:
            if outline_level and 1 <= outline_level <= 9:
                pass  # 标题是否入 flow 在下方 outline 判定后统一处理
            elif len(stripped) > 2 and not any(
                kw in stripped for kw in _BODY_SKIP_KEYWORDS
            ):
                full_flow.append(
                    _flow_item(stripped, cleaner, "paragraph", parent_num=current_num)
                )

        # ---------- 大纲提取 ----------
        if outline_level is None or outline_level < 1 or outline_level > max_level:
            continue

        title = stripped
        num_clean = ""
        if not auto_numbering:
            manual_match = _MANUAL_NUM_RE.match(title)
            if manual_match:
                num_candidate = manual_match.group(1)
                remainder = title[len(num_candidate):].lstrip(" .．、)）:：")
                if remainder and remainder[0].isalpha() and "." not in num_candidate:
                    num_clean = ""
                else:
                    num_clean = num_candidate
                    title = title[len(manual_match.group(0)):].strip()

        if not num_clean:
            auto_counters[outline_level] += 1
            for i in range(outline_level + 1, 10):
                auto_counters[i] = 0
            num_clean = ".".join(str(auto_counters[j]) for j in range(1, outline_level + 1))
        else:
            try:
                parts = [int(p) for p in num_clean.split(".") if p.isdigit()]
                if len(parts) == outline_level:
                    for i, p_val in enumerate(parts):
                        auto_counters[i + 1] = p_val
            except Exception:
                pass

        # [FIX] 剥离前缀 0 段，还原实际编号（0.4.1 -> 4.1）；
        # 层级同步调整为实际编号深度，保证父子链一致
        new_num, did_strip = _strip_leading_zero_parts(num_clean)
        if did_strip and new_num:
            num_clean = new_num
            outline_level = len(num_clean.split("."))

        title = cleaner.clean_title_suffix(title)
        if _is_body_paragraph(title):
            # 误标大纲级别的正文段落：作为正文进入全文流
            if include_full_text and len(stripped) > 2 and not any(
                kw in stripped for kw in _BODY_SKIP_KEYWORDS
            ):
                full_flow.append(
                    _flow_item(stripped, cleaner, "paragraph", parent_num=current_num)
                )
            continue

        outline.append((num_clean, title, outline_level))
        if include_full_text:
            current_num = num_clean
            full_flow.append(_flow_item(stripped, cleaner, "heading", number=num_clean))

    return outline, (full_flow if include_full_text else None)


def extract_full_text_flow(docx_path, word_items, cleaner, progress_cb=None):
    """
    供外部已有标题项（如 txt 导入）使用：按文本对齐提取全文流，
    输出与 HierarchicalMatcher._read_word_full_text 同构。
    """
    headings_queue = _build_headings_queue(word_items, cleaner)
    return _flow_pass(docx_path, headings_queue, cleaner, progress_cb=progress_cb)


def parse_tree_txt(txt_path):
    """
    反解析 Word结构树.txt（"{缩进}[编号] 标题" 行）为 outline 列表。
    层级优先按编号点数推断，无编号时按缩进（2空格一级）推断。
    """
    outline = []
    with open(txt_path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            s = line.rstrip("\r\n")
            if not s.strip():
                continue
            m = re.match(r"^(\s*)\[(.*?)\]\s?(.*)$", s)
            if not m:
                continue
            indent, num, title = m.group(1), m.group(2).strip(), m.group(3).strip()
            parts = [p for p in num.split(".") if p.strip()]
            if parts:
                level = len(parts)
            else:
                level = min(9, len(indent) // 2 + 1)
            outline.append((num, title, level))
    return outline
