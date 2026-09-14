"""
数据属性重复检测模块
基于 data_replace.py 的核心逻辑，用于检测Excel中的单功能过程内重复和跨功能过程重复
"""

import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from collections import defaultdict
import os
import datetime
from utils.runtime_logger import log_error, log_info


def clean_k(val):
    """清洗K值：统一中英文逗号，去除多余空格，保持业务内容顺序"""
    if not val:
        return ""
    return str(val).replace('，', ',').replace(' ', '').replace('\u3000', '')


def check_data_attribute_duplicates(file_path, sheet_name, keywords=None, progress_callback=None, report_base_name=None):
    """
    检测Excel中的数据属性重复
    """
    if keywords is None:
        keywords = ["ERX", "EW", "EX"]

    result = {
        'success': False,
        'output_path': None,
        'report_path': None,
        'statistics': {}
    }

    try:
        # 【修改】透传 report_base_name
        output_path, stats = _process_duplicate_detection(
            file_path, sheet_name, keywords, progress_callback,
            report_base_name=report_base_name,  # <--- 确保这里也传递了
        )

        result['success'] = True
        result['output_path'] = output_path
        result['report_path'] = output_path
        result['statistics'] = stats

    except Exception as e:
        import traceback
        result['error'] = str(e)
        log_error(f'❌ 数据属性重复检测失败: {e}')
        log_error(traceback.format_exc())

    return result


def _process_duplicate_detection(file_path, sheet_name, keywords, progress_callback, report_base_name=None):
    """
    核心处理逻辑：检测数据属性重复并生成报告

    Returns:
        tuple: (output_path, statistics_dict)
    """
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        log_info(msg)

    def report_progress(current, total, msg):
        if progress_callback:
            try:
                progress_callback(current, total, msg)
            except:
                pass

    log(f"--- 开始数据属性重复检测: {datetime.datetime.now()} ---")
    wb = openpyxl.load_workbook(file_path)

    # 如果sheet_name为None，自动使用第一个工作表
    if sheet_name is None or sheet_name not in wb.sheetnames:
        if wb.sheetnames:
            original_sheet = sheet_name
            sheet_name = wb.sheetnames[0]
            if original_sheet is None:
                log(f"[INFO] 未指定工作表，自动使用第一个工作表: '{sheet_name}'")
            else:
                log(f"[WARN] 指定的工作表 '{original_sheet}' 不存在，改用第一个工作表: '{sheet_name}'")
        else:
            raise ValueError("Excel文件中没有工作表")
    else:
        log(f"[INFO] 使用指定的工作表: '{sheet_name}'")

    ws = wb[sheet_name]
    max_row = ws.max_row

    COL_G, COL_I, COL_K, COL_Q, COL_R = 7, 9, 11, 17, 18  # Q列=17, R列=18

    # ============================================================
    # 清空K列所有颜色、条件格式和Q、R列内容
    # ============================================================
    report_progress(0, max_row - 1, "正在清空历史标记...")
    log("\n🧹 清空K列颜色、条件格式和Q、R列内容...")

    # 清空K列的条件格式
    try:
        cf_count = len(ws.conditional_formatting._cf_rules)
        ws.conditional_formatting._cf_rules = {}
        if cf_count > 0:
            log(f"  ✅ 清空了 {cf_count} 个条件格式规则")
    except:
        pass

    cleared_count = 0
    for r in range(2, max_row + 1):
        k_cell = ws.cell(r, COL_K)
        q_cell = ws.cell(r, COL_Q)
        r_cell = ws.cell(r, COL_R)

        # 清空K列填充色
        k_cell.fill = PatternFill(fill_type=None)
        cleared_count += 1

        # 清空Q、R列内容
        q_cell.value = None
        r_cell.value = None

    log(f"  ✅ 清空了 {cleared_count} 个单元格的标记")

    report_progress(0, max_row - 1, "正在划分功能块...")

    # 1. 划分连续功能块
    blocks = []
    cur, last_g = None, ""
    for r in range(2, max_row + 1):
        g = str(ws.cell(r, COL_G).value or "").strip()
        i = str(ws.cell(r, COL_I).value or "").strip().upper()
        k = str(ws.cell(r, COL_K).value or "").strip()

        if g and g != last_g:
            if cur and cur['items']:
                blocks.append(cur)
            cur = {'g': g, 'row': r, 'items': []}
            last_g = g
        if cur:
            cur['items'].append({'r': r, 't': i, 'k': k})
    if cur and cur['items']:
        blocks.append(cur)

    log(f"✅ 识别到 {len(blocks)} 个独立功能块")

    rows_red = set()
    intra_notes = {}  # 行号 -> set of note strings (单功能过程)
    cross_notes = {}  # 行号 -> set of note strings (跨功能过程)

    # 报告数据：用于生成总结报告
    intra_report = []  # 单功能过程内重复
    cross_report = []  # 跨功能过程重复

    # ============================================================
    # 2. 单功能过程内重复检测
    # ============================================================
    log("\n📊 开始单功能过程内重复检测...")
    report_progress(10, max_row - 1, "正在检测单功能过程内重复...")
    intra_total_count = 0

    for blk_idx, blk in enumerate(blocks):
        type_groups = defaultdict(list)
        for item in blk['items']:
            t = item['t']
            if t:
                type_groups[t].append(item)

        for t, items in type_groups.items():
            if len(items) < 2:
                continue

            k_groups = defaultdict(list)
            for item in items:
                ck = clean_k(item['k'])
                if ck:
                    k_groups[ck].append(item['r'])

            for ck, row_list in k_groups.items():
                if len(row_list) >= 2:
                    first_row = row_list[0]
                    dup_rows = row_list[1:]

                    rows_red.add(first_row)

                    for dup_row in dup_rows:
                        rows_red.add(dup_row)
                        intra_total_count += 1

                        if dup_row not in intra_notes:
                            intra_notes[dup_row] = set()
                        intra_notes[dup_row].add(f"N, {first_row}")

                    intra_report.append({
                        'first_row': first_row,
                        'dup_rows': dup_rows,
                        'type': t,
                        'k_value': ck,
                        'func_name': blk['g']
                    })

    log(f"  ✅ 单功能过程内发现 {len(intra_report)} 组重复，{intra_total_count} 个重复单元格")

    # ============================================================
    # 3. 跨功能过程重复检测
    # ============================================================
    log("\n开始跨功能过程模式比对...")
    report_progress(50, max_row - 1, "正在检测跨功能过程重复...")

    clean_kws = []
    for kw in keywords:
        k = kw.strip().upper()
        if k:
            clean_kws.append(k)
    log(f"  启用模式: {', '.join(clean_kws)}")

    match_groups = defaultdict(list)

    for blk in blocks:
        blk_types = {item['t'] for item in blk['items']}

        for kw_str in clean_kws:
            kw_chars = list(kw_str)
            kw_set = set(kw_chars)

            if kw_set.issubset(blk_types):
                fp_parts = []
                valid = True
                for char in kw_chars:
                    matches = [clean_k(it['k']) for it in blk['items'] if it['t'] == char]
                    if matches:
                        fp_parts.append(matches[0])
                    else:
                        valid = False
                        break
                if valid:
                    match_groups[(kw_str, tuple(fp_parts))].append(blk)

    marked_count = 0
    cross_total_count = 0

    for (pat_str, fp), group_blocks in match_groups.items():
        if len(group_blocks) < 2:
            continue

        group_blocks.sort(key=lambda b: b['row'])
        anchor_blk = group_blocks[0]

        anchor_map = {}
        for item in anchor_blk['items']:
            if item['t'] in pat_str and item['t'] not in anchor_map:
                anchor_map[item['t']] = item['r']

        log(f"  ✔️ 模式[{pat_str}] 匹配成功 | 锚点行:{anchor_blk['row']} | 涉及块数:{len(group_blocks)}")
        marked_count += 1

        for item in anchor_blk['items']:
            if item['t'] in pat_str:
                rows_red.add(item['r'])

        anchor_rows = [item['r'] for item in anchor_blk['items'] if item['t'] in pat_str]

        all_block_rows = []
        for blk in group_blocks[1:]:
            blk_rows = []
            for item in blk['items']:
                if item['t'] in pat_str:
                    rows_red.add(item['r'])
                    cross_total_count += 1
                    blk_rows.append(item['r'])

                    ref = anchor_map.get(item['t'])
                    if ref:
                        if item['r'] not in cross_notes:
                            cross_notes[item['r']] = set()
                        cross_notes[item['r']].add(f"N, {ref}")

            if blk_rows:
                all_block_rows.append(blk_rows)

        if all_block_rows and fp:
            cross_report.append({
                'anchor_rows': anchor_rows,
                'block_rows': all_block_rows,
                'pattern': pat_str,
                'k_value': fp[0] if fp else '',
                'first_func': anchor_blk['g']
            })

    log(f"\n🏁 跨块比对完成，共 {marked_count} 组模式匹配，{cross_total_count} 个重复单元格")
    report_progress(80, max_row - 1, "正在写入标记和报告...")

    # ============================================================
    # 4. 写入Q列和R列表头
    # ============================================================
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True, size=10)

    q_header = ws.cell(1, COL_Q, "单功能过程内重复")
    q_header.fill = header_fill
    q_header.font = header_font
    q_header.alignment = Alignment(horizontal='center', vertical='center')

    r_header = ws.cell(1, COL_R, "跨功能过程重复")
    r_header.fill = header_fill
    r_header.font = header_font
    r_header.alignment = Alignment(horizontal='center', vertical='center')

    # ============================================================
    # 5. 写入样式
    # ============================================================
    light_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    for r in rows_red:
        ws.cell(r, COL_K).fill = light_red

        if r in intra_notes:
            c = ws.cell(r, COL_Q)
            notes = sorted(intra_notes[r])
            c.value = "; ".join(notes)
            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        if r in cross_notes:
            c = ws.cell(r, COL_R)
            notes = sorted(cross_notes[r])
            c.value = "; ".join(notes)
            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # ============================================================
    # 6. 生成重复检测报告 Sheet
    # ============================================================
    report_sheet_name = "重复检测报告"
    if report_sheet_name in wb.sheetnames:
        del wb[report_sheet_name]
    ws_report = wb.create_sheet(report_sheet_name)

    report_headers = ["序号", "数据移动类型", "重复的K值（数据属性）", "首次出现行号",
                      "所属功能过程", "重复行号", "重复类型"]
    for col_idx, header in enumerate(report_headers, 1):
        cell = ws_report.cell(1, col_idx, header)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')

    row_idx = 1
    for rd in intra_report:
        row_idx += 1
        ws_report.cell(row_idx, 1, row_idx - 1)
        ws_report.cell(row_idx, 2, rd['type'])
        ws_report.cell(row_idx, 3, rd['k_value'])
        ws_report.cell(row_idx, 4, rd['first_row'])
        ws_report.cell(row_idx, 5, rd['func_name'])

        if rd['dup_rows']:
            dup_str = ", ".join(str(r) for r in rd['dup_rows'])
            ws_report.cell(row_idx, 6, dup_str)

        ws_report.cell(row_idx, 7, "单功能过程内重复")

    for rd in cross_report:
        row_idx += 1
        ws_report.cell(row_idx, 1, row_idx - 1)
        ws_report.cell(row_idx, 2, rd['pattern'])
        ws_report.cell(row_idx, 3, rd['k_value'])
        ws_report.cell(row_idx, 4, rd['anchor_rows'][0] if rd['anchor_rows'] else '')
        ws_report.cell(row_idx, 5, rd['first_func'])

        if rd['block_rows']:
            anchor_str = ", ".join(str(r) for r in rd['anchor_rows'])
            block_strs = [", ".join(str(r) for r in blk) for blk in rd['block_rows']]
            dup_str = anchor_str + "与" + " | ".join(block_strs)
            ws_report.cell(row_idx, 6, dup_str)

        ws_report.cell(row_idx, 7, "跨功能过程重复")

    col_widths = [8, 12, 50, 12, 30, 40, 18]
    for i, w in enumerate(col_widths, 1):
        ws_report.column_dimensions[get_column_letter(i)].width = w

    if row_idx == 1:
        ws_report.cell(2, 1, "—")
        ws_report.cell(2, 2, "—")
        ws_report.cell(2, 3, "未发现任何重复")
        ws_report.cell(2, 4, "—")
        ws_report.cell(2, 5, "—")
        ws_report.cell(2, 6, "—")
        ws_report.cell(2, 7, "—")

    # ============================================================
    # 7. 生成纯文字总结报告 Sheet
    # ============================================================
    summary_sheet_name = "总结报告"
    if summary_sheet_name in wb.sheetnames:
        del wb[summary_sheet_name]
    ws_summary = wb.create_sheet(summary_sheet_name)

    ws_summary.column_dimensions['A'].width = 120

    row = 1

    ws_summary.cell(row, 1).value = "数据处理总结报告"
    ws_summary.cell(row, 1).font = Font(bold=True, size=14)
    row += 2

    ws_summary.cell(row, 1).value = f"处理时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    row += 1
    ws_summary.cell(row, 1).value = f"源文件：{os.path.basename(file_path)}"
    row += 1
    ws_summary.cell(row, 1).value = f"工作表：{sheet_name}"
    row += 1
    ws_summary.cell(row, 1).value = f"总行数：{max_row - 1}"
    row += 1
    ws_summary.cell(row, 1).value = f"识别功能块数量：{len(blocks)}"
    row += 1
    ws_summary.cell(row, 1).value = f"启用关键词模式：{', '.join(clean_kws)}"
    row += 2

    ws_summary.cell(row, 1).value = f"标记单元格总数：{len(rows_red)}"
    row += 1
    ws_summary.cell(row, 1).value = f"单功能过程内重复组数：{len(intra_report)}"
    row += 1
    ws_summary.cell(row, 1).value = f"跨功能过程重复组数：{len(cross_report)}"
    row += 3

    ws_summary.cell(row, 1).value = "【单功能过程内重复】"
    ws_summary.cell(row, 1).font = Font(bold=True, size=12)
    row += 1

    if intra_report:
        func_groups = defaultdict(list)
        for rd in intra_report:
            func_groups[rd['func_name']].append(rd)

        for func_name, items in func_groups.items():
            ws_summary.cell(row, 1).value = f"功能过程：{func_name}"
            ws_summary.cell(row, 1).font = Font(bold=True)
            row += 1

            type_groups = defaultdict(list)
            for item in items:
                type_groups[item['type']].append(item)

            for t, t_items in type_groups.items():
                for rd in t_items:
                    first = rd['first_row']
                    dups = ", ".join(str(r) for r in rd['dup_rows'])
                    line = f"  类型{t}：{first}与{dups}（K值：{rd['k_value']}）"
                    ws_summary.cell(row, 1).value = line
                    row += 1
            row += 1
    else:
        ws_summary.cell(row, 1).value = "  未发现单功能过程内重复"
        row += 1

    row += 1

    ws_summary.cell(row, 1).value = "【跨功能过程重复】"
    ws_summary.cell(row, 1).font = Font(bold=True, size=12)
    row += 1

    if cross_report:
        pattern_groups = defaultdict(list)
        for rd in cross_report:
            pattern_groups[rd['pattern']].append(rd)

        for pattern, items in pattern_groups.items():
            ws_summary.cell(row, 1).value = f"模式：{pattern}"
            ws_summary.cell(row, 1).font = Font(bold=True)
            row += 1

            for rd in items:
                anchor_str = ", ".join(str(r) for r in rd['anchor_rows'])
                block_strs = [", ".join(str(r) for r in blk) for blk in rd['block_rows']]
                dup_str = " | ".join(block_strs)
                line = f"  {anchor_str}与{dup_str}（K值：{rd['k_value']}，功能：{rd['first_func']}）"
                ws_summary.cell(row, 1).value = line
                row += 1
            row += 1
    else:
        ws_summary.cell(row, 1).value = "  未发现跨功能过程重复"
        row += 1

    # 设置所有单元格的垂直对齐
    for cell in ws_summary['A']:
        if cell.value:
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # ============================================================
    # 8. 保存与日志 (【核心修复区：确保缩进正确】)
    # ============================================================
    from utils.initial_review_report import report_file_path

    ext = os.path.splitext(file_path)[1]

    # 使用统一的初评报告路径模块，确保进入本次运行的时间戳文件夹
    _, new_path = report_file_path(
        report_type="数据属性重复检测",
        project_name_or_path=file_path,
        report_base_name=report_base_name,
        ext=ext
    )

    wb.save(new_path)
    wb.close()

    # 合并进项目主运行日志
    from utils.runtime_logger import RuntimeLogger
    for _line in log_lines:
        if _line and str(_line).strip():
            try:
                RuntimeLogger.log(str(_line))
            except Exception:
                pass
    log_path = RuntimeLogger.get_current_log_path() or ""

    report_progress(100, max_row - 1, "检测完成")

    statistics = {
        'total_marked_rows': len(rows_red),
        'intra_duplicates': len(intra_report),
        'cross_duplicates': len(cross_report),
        'intra_total_cells': intra_total_count,
        'cross_total_cells': cross_total_count,
        'patterns_matched': marked_count,
        'blocks_identified': len(blocks),
        'output_file': new_path,
        'log_file': log_path
    }

    log(f"\n✅ 处理完成！结果文件: {new_path}")
    return new_path, statistics