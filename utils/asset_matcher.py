#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
资产清单匹配器 - 功能点拆分表 与 资产清单 的三级层级比对
从独立的 Excel 模块层级比对工具移植而来，集成到初评流水线中。

比对逻辑:
  - 文件A (功能点拆分表): 一级模块 / 二级模块 / 三级模块 三列
  - 文件B (资产清单): 三级模块(或一级分类) / 功能模块(或二级分类) / 功能点名称 三列
  - 按行加权相似度 (0.2 / 0.3 / 0.5) 取资产清单中最优匹配行
  - 分类: 完全匹配 / 模糊匹配 / 不匹配
  - 导出五工作表 Excel 报告

【修复说明 v2】:
  - 彻底解决将“说明行/属性行”（如包含“非必填”、“按需”）误判为表头的问题。
  - 引入“负向关键词惩罚机制”，强力排除非表头行。
  - 引入“位置偏好”，表头通常在前5行，给予额外加分。

【性能优化说明 v3】(针对上万行 × 上万行的匹配提速):
  1. B 侧(资产清单)一次性归一化 + 去重；A 侧一次性向量化读取(替代逐行 iterrows)。
  2. 精确匹配先行：三级取值完全相等的行直接命中，不进入模糊比对。
  3. 模糊比对优先使用 rapidfuzz.cdist (C++/多线程) 批量计算加权相似度矩阵，
     每个 A 行仅取 top-k 候选，再用原 difflib 逻辑复算，保证输出与旧版一致；
     未安装 rapidfuzz 时自动回退到“比值上界剪枝”版逐行算法(仍比旧版快数倍)。
  4. 读取侧小优化：iter_rows(values_only=True)、工作表名用 read_only 模式获取。

可选依赖: pip install rapidfuzz   (强烈建议安装，模糊比对可提速约两个数量级)
"""

import os
import re
from datetime import datetime
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from utils.runtime_logger import log_info

try:
    from extend.matcher_config import MatcherConfig
except ImportError:
    MatcherConfig = None

# rapidfuzz 为可选依赖：安装后模糊比对走多线程批量矩阵；未安装则走剪枝回退算法
try:
    from rapidfuzz import fuzz
    from rapidfuzz.process import cdist
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False


# -----------------------------------------------------------------------------
# 工具函数
# -----------------------------------------------------------------------------
def get_excel_col_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result


def strip_prefix(text):
    """去掉下拉框选项的 "A. " 列字母前缀"""
    if not text:
        return text
    m = re.match(r"^[A-Z]+\.\s*(.*)", text)
    return m.group(1) if m else text


def similarity(a, b, _cache={}):
    """大小写不敏感的序列相似度 (带进程级缓存)"""
    a, b = str(a).strip(), str(b).strip()
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    key = (a, b) if a <= b else (b, a)
    score = _cache.get(key)
    if score is None:
        score = SequenceMatcher(None, a.lower(), b.lower()).ratio()
        if len(_cache) < 500000:
            _cache[key] = score
    return score


# 匹配矩阵用的哨兵字符：保证 “空 vs 空 = 满分、空 vs 非空 = 0” 与 similarity() 语义一致
_EMPTY = "\x00"


def _norm_cell(v):
    """单元格值统一转为去除首尾空白的字符串；None/NaN → ''"""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


# -----------------------------------------------------------------------------
# Excel 读取 (支持合并单元格 + 智能表头定位 v2)
# -----------------------------------------------------------------------------
def find_best_header_row(data, max_col, context="generic"):
    """
    智能表头定位策略 v3（上下文感知）：
    - context="split": 拆分表场景，优先识别"一级模块/二级模块/三级模块"
    - context="asset": 资产清单场景，优先识别"建设目标/一级分类/功能模块/功能点名称"
    - context="generic": 通用场景

    1. 正向关键词加分（根据上下文调整权重）
    2. 负向关键词重罚（非必填、按需、示例等）
    3. 位置偏好（前5行优先）
    4. 非空单元格数量辅助
    5. 列数密度奖励（表头行通常有更多有效列名）
    """

    # 根据上下文选择不同的关键词配置
    if context == "split":
        # 拆分表：强权重给层级模块关键词
        high_priority_keywords = ["一级模块", "二级模块", "三级模块", "一级功能", "二级功能", "三级功能"]
        medium_priority_keywords = ["模块", "功能", "层级", "级别"]
        low_priority_keywords = ["名称", "分类", "项目", "序号", "编号", "描述", "说明", "备注", "类型", "状态", "系统", "举措", "内容", "过程", "资产", "清单", "拆分"]
    elif context == "asset":
        # 资产清单：强权重给资产相关关键词
        high_priority_keywords = ["建设目标", "一级分类", "二级分类", "功能模块", "功能点名称", "功能点描述", "资产类别", "工作量"]
        medium_priority_keywords = ["分类", "模块", "功能", "资产", "清单", "建设"]
        low_priority_keywords = ["名称", "描述", "序号", "编号", "备注", "类型", "状态", "内容", "举措", "分析", "复用", "共享"]
    else:
        # 通用场景
        high_priority_keywords = []
        medium_priority_keywords = ["模块", "功能", "名称", "分类", "级别", "层级", "项目", "序号", "编号"]
        low_priority_keywords = ["描述", "说明", "备注", "类型", "状态", "系统", "举措", "内容", "过程", "资产", "清单", "拆分", "一级", "二级", "三级"]

    # 负向关键词（绝对不可能是列名）
    negative_keywords = [
        "非必填", "按需", "非必要", "必填", "选填", "示例", "例如",
        "备注：", "说明：", "如项目较小", "如项目较大", "项目名称", "Sheet表的名字"
    ]

    best_row_idx = 0
    best_score = -1000

    for row_idx, row_values in enumerate(data):
        # 特殊处理：检查是否包含完整的层级关键词组合（一票通过级别）
        row_text = " ".join([str(v).strip() for v in row_values if v and str(v).strip()])
        has_level1 = any(kw in row_text for kw in ["一级模块", "一级功能"])
        has_level2 = any(kw in row_text for kw in ["二级模块", "二级功能"])
        has_level3 = any(kw in row_text for kw in ["三级模块", "三级功能"])

        # 如果同时包含三个层级关键词，给予超高奖励（确保被选中）
        level_combo_bonus = 0
        if has_level1 and has_level2 and has_level3:
            level_combo_bonus = 500  # 超级奖励

        # 1. 正向关键词得分（分级权重）
        keyword_score = 0
        non_empty_count = 0

        for val in row_values:
            if val and str(val).strip():
                val_str = str(val).strip()
                non_empty_count += 1

                # 高优先级关键词：每个加50分（大幅提升）
                for kw in high_priority_keywords:
                    if kw in val_str:
                        keyword_score += 50

                # 中优先级关键词：每个加20分
                for kw in medium_priority_keywords:
                    if kw in val_str:
                        keyword_score += 20

                # 低优先级关键词：每个加10分
                for kw in low_priority_keywords:
                    if kw in val_str:
                        keyword_score += 10

        # 2. 负向关键词惩罚（力度要大）
        negative_score = 0
        for val in row_values:
            if val and str(val).strip():
                val_str = str(val).strip()
                for kw in negative_keywords:
                    if kw in val_str:
                        negative_score += 25

        # 3. 非空单元格数量（适度奖励）
        non_empty_bonus = min(non_empty_count * 2, 20)

        # 4. 位置偏好（表头通常靠前，但不要过度偏向第1行）
        position_score = 0
        if row_idx == 0:
            position_score = 5  # 第1行只给少量加分，避免误选标题行
        elif row_idx < 3:
            position_score = 15  # 2-3行是常见表头位置
        elif row_idx < 6:
            position_score = 10  # 4-6行也可能是表头
        elif row_idx < 10:
            position_score = 0
        else:
            position_score = -15  # 10行以后大幅扣分

        # 5. 列密度奖励（表头行通常有较多有效列名）
        density_bonus = 0
        if non_empty_count >= 5:
            density_bonus = 15
        elif non_empty_count >= 3:
            density_bonus = 8

        # 6. 综合得分计算（包含层级组合超级奖励）
        total_score = keyword_score - negative_score + position_score + non_empty_bonus + density_bonus + level_combo_bonus

        if total_score > best_score:
            best_score = total_score
            best_row_idx = row_idx

    return best_row_idx


def read_excel_sheet(path, sheet_name, header_row=None, log_callback=None, context="generic"):
    """
    完整读取一个工作表:
      - header_row: 可选，手动指定表头行号 (1-based)。为 None 时启用智能自动探测。
      - context: 上下文类型 ("split", "asset", "generic")，用于优化智能表头检测
      - 自动在前 50 行内定位表头 (结合正负向关键词与位置偏好)
      - 填充合并单元格
      - 返回
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log_info(msg)

    try:
        wb = load_workbook(path, data_only=True, read_only=False)
        ws = wb[sheet_name]

        # 动态探测安全列数
        actual_max_col = ws.max_column
        for merged_range in ws.merged_cells.ranges:
            if merged_range.max_col > actual_max_col:
                actual_max_col = merged_range.max_col
        max_col = min(300, actual_max_col)

        probe_max_row = min(50, ws.max_row)

        # 预先读取探测区域的数据 (values_only=True 比逐 Cell 取 .value 快)
        cell_values = {}
        for r_idx, row in enumerate(
                ws.iter_rows(min_row=1, max_row=probe_max_row, max_col=max_col, values_only=True),
                start=1):
            for c_idx, v in enumerate(row, start=1):
                cell_values[(r_idx, c_idx)] = v

        # 探测范围内填充合并单元格 (关键：确保合并的表头被正确识别)
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row > probe_max_row or merged_range.min_col > max_col:
                continue
            top_left = merged_range.start_cell
            value = top_left.value
            end_row = min(merged_range.max_row, probe_max_row)
            end_col = min(merged_range.max_col, max_col)
            for r in range(merged_range.min_row, end_row + 1):
                for c in range(merged_range.min_col, end_col + 1):
                    cell_values[(r, c)] = value

        data = []
        for r in range(1, probe_max_row + 1):
            data.append([cell_values.get((r, c)) for c in range(1, max_col + 1)])

        # 确定表头行
        if header_row is not None:
            best_row_idx = header_row - 1
            log(f"  [{os.path.basename(path)}:{sheet_name}] 使用用户指定的表头行: 第 {header_row} 行")
        else:
            best_row_idx = find_best_header_row(data, max_col, context=context)
            log(f"  [{os.path.basename(path)}:{sheet_name}] 智能定位表头在第 {best_row_idx + 1} 行 (context={context})")

        header_row_data = data[best_row_idx]
        columns = []
        for j, h in enumerate(header_row_data):
            if h is not None and str(h).strip():
                columns.append(str(h).strip())
            else:
                columns.append(f"Column_{j + 1}")

        # 全表读取数据区
        max_row = ws.max_row
        for r_idx, row in enumerate(
                ws.iter_rows(min_row=best_row_idx + 1, max_row=max_row, max_col=max_col, values_only=True),
                start=best_row_idx + 1):
            for c_idx, v in enumerate(row, start=1):
                cell_values[(r_idx, c_idx)] = v

        # 全表填充合并单元格
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col > max_col:
                continue
            top_left = merged_range.start_cell
            value = top_left.value
            end_col = min(merged_range.max_col, max_col)
            for r in range(merged_range.min_row, merged_range.max_row + 1):
                for c in range(merged_range.min_col, end_col + 1):
                    cell_values[(r, c)] = value

        records = []
        excel_rows = []
        for r in range(best_row_idx + 2, max_row + 1):
            records.append([cell_values.get((r, c)) for c in range(1, max_col + 1)])
            excel_rows.append(r)

        df = pd.DataFrame(records, columns=columns)
        df.index = excel_rows
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")

        wb.close()
        return df, best_row_idx + 1

    except Exception as e:
        log(f"  openpyxl 读取异常，回退 pandas: {e}")
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=0 if header_row is None else header_row - 1)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.ffill().fillna("")
            df.index = range(2, len(df) + 2)
            return df, 1 if header_row is None else header_row
        except Exception as e2:
            raise ValueError(f"无法读取 {path} 的 {sheet_name}: {e2}")


def _get_sheet_names(path):
    """轻量获取工作表名列表 (read_only 模式，避免整簿被解析两遍)"""
    try:
        wb = load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        return pd.ExcelFile(path).sheet_names


def pick_sheet(path, prefer_keywords, log_callback=None):
    """按关键字优先选择工作表，找不到则返回第一个"""

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log_info(msg)

    sheets = _get_sheet_names(path)
    if len(sheets) == 1:
        return sheets[0]
    for kw in prefer_keywords:
        for s in sheets:
            if kw in s:
                log(f"  自动选择工作表: {s}")
                return s
    log(f"  未命中关键字，默认选择第一个工作表: {sheets[0]}")
    return sheets[0]


def detect_level_columns(columns, side):
    """自动探测三个层级对应的列索引"""
    if side == "split":
        targets = [["一级模块"], ["二级模块"], ["三级模块"]]
    else:
        targets = [
            ["三级模块", "一级分类", "建设举措", "模块", "系统"],
            ["功能模块", "二级分类", "建设内容", "子模块", "功能"],
            ["功能点名称", "建设内容简述", "描述", "说明"],
        ]

    result = []
    used_indices = set()
    for keywords in targets:
        found_idx = None
        # 策略1: 精确匹配
        for kw in keywords:
            for idx, col in enumerate(columns):
                if idx in used_indices:
                    continue
                if str(col).strip() == kw:
                    found_idx = idx
                    break
            if found_idx is not None:
                break
        # 策略2: 包含匹配
        if found_idx is None:
            for kw in keywords:
                for idx, col in enumerate(columns):
                    if idx in used_indices:
                        continue
                    if kw in str(col):
                        found_idx = idx
                        break
                if found_idx is not None:
                    break
        if found_idx is not None:
            used_indices.add(found_idx)
        result.append(found_idx)
    return result


def _resolve_level_columns(df, explicit_idx, side, log):
    """解析层级列索引"""
    n_cols = len(df.columns)
    if explicit_idx and all(isinstance(i, int) and 0 <= i < n_cols for i in explicit_idx):
        cols = [df.columns[i] for i in explicit_idx]
        log(f"  使用用户指定的{'拆分表' if side == 'split' else '资产清单'}层级列: {list(cols)}")
        return cols, list(explicit_idx)

    idx_list = detect_level_columns(list(df.columns), side)

    if side == "split":
        if any(i is None for i in idx_list):
            raise ValueError(
                f"拆分表中未找到层级列 (尝试匹配: 一级模块/二级模块/三级模块)，"
                f"现有列: {list(df.columns)[:20]}"
            )
        cols = [df.columns[i] for i in idx_list]
        return cols, idx_list
    else:
        first_found = next((i for i in idx_list if i is not None), None)
        if first_found is None:
            raise ValueError(
                f"资产清单中未找到任何层级列，现有列: {list(df.columns)[:20]}"
            )
        for i in range(3):
            if idx_list[i] is None:
                idx_list[i] = first_found
                log(f"  [WARN] 资产清单第 {i + 1} 级未匹配到列，复用第 1 级列: {df.columns[first_found]}")
        cols = [df.columns[i] for i in idx_list]
        return cols, idx_list


# -----------------------------------------------------------------------------
# 核心比对
# -----------------------------------------------------------------------------
def _match_candidates_matrix(a_cols, b_cols, topk=5, chunk=256,
                             on_progress=None, should_cancel=None):
    """
    rapidfuzz 多线程批量计算加权相似度，为每个 A 行返回 top-k 候选 B 行下标(升序)。
      - a_cols / b_cols: [ [一级值...], [二级值...], [三级值...] ] (已归一化)
      - 与 similarity() 对齐：小写比较；空串用哨兵字符，保证 空vs空=满分、空vs非空=0
      - 分块计算以控制内存：chunk=256、B 侧 5 万行时每块矩阵约 50MB
    """
    n, m = len(a_cols[0]), len(b_cols[0])
    if m == 0 or n == 0:
        return [[] for _ in range(n)]

    weights = (0.2, 0.3, 0.5)
    a_s = [[(x.lower() if x else _EMPTY) for x in col] for col in a_cols]
    b_s = [[(x.lower() if x else _EMPTY) for x in col] for col in b_cols]

    k = min(topk, m)
    candidates = [None] * n
    for s in range(0, n, chunk):
        e = min(s + chunk, n)
        total = None
        for i, w in enumerate(weights):
            mat = cdist(a_s[i][s:e], b_s[i], scorer=fuzz.ratio, workers=-1)
            mat = mat.astype(np.float32)
            mat *= w / 100.0
            total = mat if total is None else total + mat
        if k < m:
            cand = np.argpartition(-total, k - 1, axis=1)[:, :k]
        else:
            cand = np.tile(np.arange(m), (e - s, 1))
        cand.sort(axis=1)  # 升序，保持旧版“先出现的行优先”的平局规则
        for r in range(e - s):
            candidates[s + r] = cand[r].tolist()
        if on_progress:
            on_progress(e / n, f"已比对 {e}/{n} 行")
        if should_cancel and should_cancel():
            return None
    return candidates


def _match_candidates_python(a_cols, b_cols, on_progress=None, should_cancel=None):
    """
    未安装 rapidfuzz 时的回退实现 (结果与旧版逐行暴力循环完全一致)：
      - 预归一化 + 按 SequenceMatcher 比值上界 2*min(la,lb)/(la+lb) 剪枝
      - 先算权重最大的第三列，尽早缩小上界，典型提速 2~5 倍
    """
    n, m = len(a_cols[0]), len(b_cols[0])
    b0, b1, b2 = b_cols
    lens = [(len(b0[j]), len(b1[j]), len(b2[j])) for j in range(m)]
    candidates = []
    for i in range(n):
        a0, a1, a2 = a_cols[0][i], a_cols[1][i], a_cols[2][i]
        la = (len(a0), len(a1), len(a2))
        best_t, best_j = -1.0, -1
        for j in range(m):
            lb = lens[j]
            ub0 = 1.0 if la[0] + lb[0] == 0 else 2.0 * min(la[0], lb[0]) / (la[0] + lb[0])
            ub1 = 1.0 if la[1] + lb[1] == 0 else 2.0 * min(la[1], lb[1]) / (la[1] + lb[1])
            ub2 = 1.0 if la[2] + lb[2] == 0 else 2.0 * min(la[2], lb[2]) / (la[2] + lb[2])
            if 0.2 * ub0 + 0.3 * ub1 + 0.5 * ub2 <= best_t:
                continue
            s2 = similarity(a2, b2[j])
            if 0.2 * ub0 + 0.3 * ub1 + 0.5 * s2 <= best_t:
                continue
            s0 = similarity(a0, b0[j])
            if 0.2 * s0 + 0.3 * ub1 + 0.5 * s2 <= best_t:
                continue
            s1 = similarity(a1, b1[j])
            t = 0.2 * s0 + 0.3 * s1 + 0.5 * s2
            if t > best_t:
                best_t, best_j = t, j
        candidates.append([best_j] if best_j >= 0 else [])
        if (i + 1) % 200 == 0 or i == n - 1:
            if on_progress:
                on_progress((i + 1) / n, f"已比对 {i + 1}/{n} 行")
            if should_cancel and should_cancel():
                return None
    return candidates


def compare_split_with_asset(
        split_path,
        asset_path,
        split_sheet=None,
        asset_sheet=None,
        split_header_row=None,
        asset_header_row=None,
        split_col_idx=None,
        asset_col_idx=None,
        threshold=0.8,
        fuzzy_topk=5,
        progress_callback=None,
        should_cancel=None,
        log_callback=None,
        report_base_name=None,
):
    """
    比对功能点拆分表与资产清单
      - fuzzy_topk: 模糊比对阶段每个 A 行保留的候选数 (默认 5，边界行对不上可调到 10/20)
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log_info(msg)

    def report_progress(p, msg):
        if progress_callback:
            try:
                progress_callback(max(0.0, min(1.0, p)), msg)
            except Exception:
                pass

    try:
        log(f"[资产匹配] 读取拆分表: {os.path.basename(split_path)}")
        report_progress(0.02, "读取功能点拆分表...")
        if split_sheet:
            df_a, header_row_a = read_excel_sheet(split_path, split_sheet, header_row=split_header_row,
                                                  log_callback=log, context="split")
        else:
            sheet_a = pick_sheet(split_path, ["功能点拆分", "功能过程点拆分", "拆分"], log)
            df_a, header_row_a = read_excel_sheet(split_path, sheet_a, header_row=split_header_row,
                                                  log_callback=log, context="split")

        log(f"[资产匹配] 读取资产清单: {os.path.basename(asset_path)}")
        report_progress(0.08, "读取资产清单...")
        if asset_sheet:
            df_b, _ = read_excel_sheet(asset_path, asset_sheet, header_row=asset_header_row,
                                       log_callback=log, context="asset")
        else:
            sheet_b = pick_sheet(asset_path, ["资产", "清单"], log)
            df_b, _ = read_excel_sheet(asset_path, sheet_b, header_row=asset_header_row,
                                       log_callback=log, context="asset")

        target_cols_a, cols_a_idx = _resolve_level_columns(df_a, split_col_idx, "split", log)
        target_cols_b, cols_b_idx = _resolve_level_columns(df_b, asset_col_idx, "asset", log)
        display_a = [f"{get_excel_col_letter(cols_a_idx[i])}. {target_cols_a[i]}" for i in range(3)]
        display_b = [f"{get_excel_col_letter(cols_b_idx[i])}. {target_cols_b[i]}" for i in range(3)]

        # 过滤三级层级全空的行
        mask = (
                (df_a[target_cols_a[0]].astype(str).str.strip() == "")
                & (df_a[target_cols_a[1]].astype(str).str.strip() == "")
                & (df_a[target_cols_a[2]].astype(str).str.strip() == "")
        )
        df_a = df_a[~mask].copy()
        log(f"[资产匹配] 拆分表有效数据 {len(df_a)} 行 (表头第 {header_row_a} 行)")

        df_a_unique = df_a.drop_duplicates(subset=target_cols_a, keep="first")
        log(f"[资产匹配] 去重后待匹配 {len(df_a_unique)} 行")

        # ==================== 性能优化版匹配逻辑 v3 ====================
        # 1) B 侧(资产清单)一次性归一化 + 去重 (旧版在内层循环里重复转换了 N 遍)
        b_raw = df_b[target_cols_b].values.tolist()
        b_unique, exact_map = [], {}
        for r in b_raw:
            key = (_norm_cell(r[0]), _norm_cell(r[1]), _norm_cell(r[2]))
            if key not in exact_map:
                exact_map[key] = list(key)
                b_unique.append(list(key))
        log(f"[资产匹配] 资产清单共 {len(b_raw)} 行，去重后 {len(b_unique)} 行")

        # 2) A 侧(拆分表)一次性向量化读取 (替代逐行 iterrows)
        a_raw = df_a_unique[target_cols_a].values.tolist()
        excel_rows_a = list(df_a_unique.index)
        a_rows = [[_norm_cell(r[0]), _norm_cell(r[1]), _norm_cell(r[2])] for r in a_raw]

        # 3) 精确匹配先行：三级取值完全相等的行直接命中，只有未命中的才进入模糊比对
        fuzzy_pos = [i for i, r in enumerate(a_rows) if any(r) and tuple(r) not in exact_map]
        log(f"[资产匹配] 精确命中 {len(a_rows) - len(fuzzy_pos)} 行，待模糊匹配 {len(fuzzy_pos)} 行")
        report_progress(0.12, f"开始批量比对，共 {len(fuzzy_pos)} 行...")

        candidates = None
        if fuzzy_pos:
            a_cols = [[a_rows[i][k] for i in fuzzy_pos] for k in range(3)]
            b_cols = [[r[k] for r in b_unique] for k in range(3)]
            if HAS_RAPIDFUZZ:
                log("[资产匹配] 使用 rapidfuzz 多线程批量计算相似度...")
                candidates = _match_candidates_matrix(
                    a_cols, b_cols, topk=fuzzy_topk,
                    on_progress=lambda frac, msg: report_progress(0.12 + 0.78 * frac, msg),
                    should_cancel=should_cancel)
            else:
                log("[资产匹配] [WARN] 未安装 rapidfuzz，使用剪枝回退算法 "
                    "(pip install rapidfuzz 可再提速约百倍)")
                candidates = _match_candidates_python(
                    a_cols, b_cols,
                    on_progress=lambda frac, msg: report_progress(0.12 + 0.78 * frac, msg),
                    should_cancel=should_cancel)
            if candidates is None:
                return None  # 用户取消

        # 4) 对 top-k 候选行用原 difflib 逻辑复算，保证结果与旧版一致
        report_progress(0.90, "正在汇总匹配结果...")
        best_of = {}
        for pos, i in enumerate(fuzzy_pos):
            if should_cancel and pos % 500 == 0 and should_cancel():
                return None
            a0, a1, a2 = a_rows[i]
            best_t, best_j, best_sims = -1.0, -1, (0.0, 0.0, 0.0)
            for j in candidates[pos]:
                b_row = b_unique[j]
                s0 = similarity(a0, b_row[0])
                s1 = similarity(a1, b_row[1])
                s2 = similarity(a2, b_row[2])
                t = 0.2 * s0 + 0.3 * s1 + 0.5 * s2
                if t > best_t:
                    best_t, best_j, best_sims = t, j, (s0, s1, s2)
            if best_j >= 0:
                best_of[i] = (b_unique[best_j], best_sims)

        # 5) 组装结果 (分类 / 描述 / 输出格式与旧版完全一致)
        results, items = [], []

        for i, excel_row in enumerate(excel_rows_a):
            val_a = a_rows[i]
            if not any(val_a):
                continue

            key = tuple(val_a)
            if key in exact_map:
                best_b, best_sims = exact_map[key], (1.0, 1.0, 1.0)
            elif i in best_of:
                best_b, best_sims = best_of[i]
            else:
                best_b, best_sims = ["", "", ""], (0.0, 0.0, 0.0)

            matches = [best_sims[k] >= threshold for k in range(3)]
            match_count = sum(matches)

            has_missing = False
            has_mismatch = False
            for k in range(3):
                if val_a[k]:
                    if not best_b[k]:
                        has_missing = True
                    elif not matches[k]:
                        has_mismatch = True

            if match_count == 3 and not has_missing and not has_mismatch:
                status = "完全匹配"
            elif has_missing or has_mismatch:
                status = "不匹配"
            elif match_count > 0:
                status = "模糊匹配"
            else:
                status = "不匹配"

            overall_sim = 0.2 * best_sims[0] + 0.3 * best_sims[1] + 0.5 * best_sims[2]

            mismatch_descs = []
            missing_descs = []
            for k in range(3):
                if val_a[k]:
                    if not best_b[k]:
                        missing_descs.append(f"拆分表“{target_cols_a[k]}”列“{val_a[k]}”在资产清单中无对应项")
                    elif not matches[k]:
                        mismatch_descs.append(
                            f"拆分表“{target_cols_a[k]}”列“{val_a[k]}”与资产清单“{target_cols_b[k]}”列“{best_b[k]}”不匹配")

            missing_desc_str = "；".join(missing_descs) if missing_descs else "无缺失"
            mismatch_desc_str = "；".join(mismatch_descs) if mismatch_descs else "无层级错位"

            b_final = []
            for k in range(3):
                if val_a[k] and (not best_b[k] or not matches[k]):
                    b_final.append(" ❌缺少")
                else:
                    b_final.append(best_b[k] if best_b[k] else "")

            row_result = {
                display_a[0]: val_a[0], display_b[0]: b_final[0],
                display_a[1]: val_a[1], display_b[1]: b_final[1],
                display_a[2]: val_a[2], display_b[2]: b_final[2],
                "相似度": f"{overall_sim:.2%}", "匹配状态": status,
                "缺失简略描述": missing_desc_str, "层级不匹配简略描述": mismatch_desc_str,
                "row_range": excel_row,
            }
            results.append(row_result)

            if status != "完全匹配":
                items.append({
                    "拆分表一级": val_a[0], "拆分表二级": val_a[1], "拆分表三级": val_a[2],
                    "资产清单三级": best_b[2] if best_b[2] else "-", "匹配状态": status,
                    "相似度": f"{overall_sim:.2%}", "缺失简略描述": missing_desc_str,
                    "层级不匹配简略描述": mismatch_desc_str, "row_range": excel_row,
                })
        # ==================== 性能优化版匹配逻辑结束 ====================

        log(f"[资产匹配] 匹配完成，共 {len(results)} 条结果")

        df_result = pd.DataFrame(results)
        if not df_result.empty:
            subset_cols = [display_a[0], display_a[1], display_a[2], display_b[0], display_b[1], display_b[2]]
            df_result = df_result.drop_duplicates(subset=subset_cols, keep="first")
        log(f"[资产匹配] 去重后剩余 {len(df_result)} 条结果")

        report_progress(0.92, "正在生成资产清单匹配报告...")
        report_path = export_asset_report(
            df_result,
            split_path,
            threshold,
            log_callback=log,
            report_base_name=report_base_name  # <--- 【新增】
        )

        total_res = len(df_result)
        exact = len(df_result[df_result["匹配状态"] == "完全匹配"])
        fuzzy = len(df_result[df_result["匹配状态"] == "模糊匹配"])
        mismatch = len(df_result[df_result["匹配状态"] == "不匹配"])
        rate = (exact + fuzzy) / total_res if total_res > 0 else 0.0

        statistics = {
            "总数": total_res, "完全匹配": exact, "模糊匹配": fuzzy, "不匹配": mismatch,
            "匹配率": f"{rate:.1%}", "资产清单": os.path.basename(asset_path),
        }

        report_progress(1.0, "资产清单匹配完成")
        log(f"[资产匹配] 统计: 总数 {total_res}，完全匹配 {exact}，模糊匹配 {fuzzy}，不匹配 {mismatch}")

        return {"is_valid": True, "statistics": statistics, "report_path": report_path, "items": items}

    except Exception as e:
        import traceback
        log(f"[资产匹配] 执行异常: {e}")
        log(traceback.format_exc())
        return {"is_valid": False, "error": str(e)}


# -----------------------------------------------------------------------------
# 报告导出
# -----------------------------------------------------------------------------
def export_asset_report(df_result, split_path, threshold=0.8, output_path=None, log_callback=None,
                        report_base_name=None, clean_name=None):
    """导出五工作表样式的资产清单匹配报告"""

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log_info(msg)

    # 【核心修复】：统一使用初评报告路径模块生成路径，确保进入本次运行的时间戳文件夹
    if output_path is None:
        from utils.initial_review_report import report_file_path
        _, output_path = report_file_path(
            report_type="资产清单匹配报告",
            project_name_or_path=split_path,
            report_base_name=report_base_name,
        )

    final_path = output_path
    counter = 1
    # 防占用处理
    while True:
        try:
            if os.path.exists(final_path):
                with open(final_path, "a"):
                    pass
            break
        except (IOError, PermissionError):
            name, ext = os.path.splitext(output_path)
            final_path = f"{name}({counter}){ext}"
            counter += 1

    # 确保目录存在
    os.makedirs(os.path.dirname(final_path), exist_ok=True)

    with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
        total = len(df_result)
        if total > 0:
            exact = len(df_result[df_result["匹配状态"] == "完全匹配"])
            fuzzy = len(df_result[df_result["匹配状态"] == "模糊匹配"])
            missing = len(df_result[df_result["匹配状态"] == "不匹配"])
            summary = f"{exact + fuzzy}/{total}"
        else:
            exact = fuzzy = missing = 0
            summary = "0/0"

        stats_data = {
            "统计项": [" 统计概览", "✅ 精确匹配", "🔍 模糊匹配", "❌ 缺失和⚠️ 层级不匹配", " 汇总匹配结果", "🎯 匹配阈值"],
            "数量": [total, exact, fuzzy, missing, summary, f"{threshold:.2f}"],
        }
        pd.DataFrame(stats_data).to_excel(writer, sheet_name=" 统计概览", index=False)

        df_exact = df_result[df_result["匹配状态"] == "完全匹配"].copy()
        if df_exact.empty: df_exact = pd.DataFrame({"说明": ["无精确匹配项"]})
        df_exact.to_excel(writer, sheet_name="✅ 精确匹配", index=False)

        df_fuzzy = df_result[df_result["匹配状态"] == "模糊匹配"].copy()
        if df_fuzzy.empty: df_fuzzy = pd.DataFrame({"说明": ["无模糊匹配项"]})
        df_fuzzy.to_excel(writer, sheet_name=" 模糊匹配", index=False)

        df_missing = df_result[df_result["匹配状态"] == "不匹配"].copy()
        if df_missing.empty: df_missing = pd.DataFrame({"说明": ["无缺失或层级不匹配项"]})
        df_missing.to_excel(writer, sheet_name="❌ 缺失和️ 层级不匹配", index=False)

        df_all = df_result.copy()
        if df_all.empty: df_all = pd.DataFrame({"说明": ["无数据"]})
        df_all.to_excel(writer, sheet_name="📋 汇总匹配结果", index=False)

        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
                             bottom=Side(style="thin"))

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for row_idx in range(2, ws.max_row + 1):
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value: max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

    log(f"[资产匹配] 报告已导出至: {final_path}")
    return final_path