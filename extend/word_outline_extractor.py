# extend/word_outline_extractor.py
"""
Word 大纲提取器（增强版）
- 严格使用 OutlineLevel 1-9 过滤（纯净大纲）
- 同时支持自动编号（ListFormat）和手动编号（正则提取）
"""
import win32com.client
import pythoncom
import os
import re  # 新增：用于手动编号提取
from typing import List, Tuple
import time  # 新增：用于性能监测

# 优先使用项目统一日志；作为独立脚本从 extend/ 目录直接运行时降级为 print
try:
    from utils.runtime_logger import log_error, log_info, log_warn
except ImportError:
    def log_info(*parts, sep=" "):
        print(sep.join(str(p) for p in parts))

    def log_warn(*parts, sep=" "):
        print(sep.join(str(p) for p in parts))

    def log_error(*parts, sep=" "):
        print(sep.join(str(p) for p in parts))


def extract_word_outline(doc_path: str, max_level: int = 9) -> List[Tuple[str, str, int]]:
    """
    稳定提取 Word 大纲（OutlineLevel 1~9），支持自动/手动编号

        返回: [(编号, 标题, 展示级别), ...]
        - 过滤口径：仍使用 Word 的 OutlineLevel 1~9 来筛选“纯净大纲”段落
        - 展示口径：为了让目录树呈现为 4.6 / 4.6.1 / 4.6.1.1 ... 的缩进结构，
            对于形如 4.6.1.1 的数字编号，展示级别按“点的个数”计算：
            4.6 -> 1，4.6.1 -> 2，4.6.1.1 -> 3 ...（上限 9）
    """
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        doc_path = os.path.abspath(doc_path)
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"文件不存在: {doc_path}")

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3

        doc = word.Documents.Open(
            FileName=doc_path,
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False
        )

        if doc is None:
            raise RuntimeError("文档对象为空")

        result = []

        for para in doc.Paragraphs:
            try:
                text = para.Range.Text
                if not text or text.strip() in ['\r', '\x07', '']:
                    continue

                outline_level = para.OutlineLevel
                if outline_level < 1 or outline_level > 9:
                    continue
                if outline_level > max_level:
                    continue

                # ===== 核心增强：双重编号提取 =====
                # 1. 优先尝试自动编号（Word 原生 ListFormat）
                list_text = para.Range.ListFormat.ListString.strip()
                title = text.strip().replace('\r', '').replace('\x07', '')

                # 2. 如果自动编号为空，尝试从文本开头提取手动编号
                if not list_text:
                    # 匹配模式: 1.1.1.1 / 4.1.1.2 / 4.1.1.2、 / 4.1.1.2 查询...
                    # 限制: 最多4级编号（避免匹配"2023年"等）
                    match = re.match(
                        r'^(\d+(?:[\.\．]\d+){0,3})[\s\.．、：:\u3000]?',  # 允许编号后无空格直接接文字
                        title
                    )
                    if match:
                        list_text = match.group(1).replace('．', '.')  # 统一为半角点
                        # 移除编号部分（包括可能的分隔符）
                        title = title[len(match.group(0)):].strip()
                else:
                    # 自动编号存在时，清理标题中重复的编号前缀
                    if title.startswith(list_text):
                        title = title[len(list_text):].strip()
                    # 处理 ListString="4.1.1." 但标题="4.1.1 查询" 的情况
                    elif list_text.endswith('.'):
                        prefix = list_text.rstrip('.')
                        if title.startswith(prefix):
                            title = title[len(prefix):].strip()

                # 清理编号末尾标点
                num_clean = list_text.rstrip('.').rstrip(')').rstrip('：').rstrip(':').strip()

                # 展示级别：优先按数字编号的“点数”映射目录层级
                display_level = outline_level
                if num_clean and re.match(r"^\d+(?:\.\d+)*$", num_clean):
                    # 修正：段数 = 点数 + 1
                    # "4" (0个点) -> 1级, "4.1" (1个点) -> 2级, "4.1.1" (2个点) -> 3级
                    display_level = min(9, max(1, num_clean.count('.') + 1))

                if title:
                    result.append([num_clean, title, display_level])

            except pythoncom.com_error:
                continue
            except Exception:
                continue

        return result

    finally:
        if doc:
            try:
                doc.Close(False)
            except:
                pass
        if word:
            try:
                word.Quit()
            except:
                pass
        pythoncom.CoUninitialize()


def print_word_outline(outline: List[Tuple[str, str, int]]) -> None:
    """简洁输出大纲结构"""
    if not outline:
        log_warn('\n⚠️ 未提取到任何大纲项')
        log_info('💡 请在 Word 中验证：【视图】→【大纲】是否显示层级结构')
        return

    log_info(f"\n{'=' * 70}")
    log_info('Word 大纲视图 (OutlineLevel 1-9)')
    log_info(f"{'=' * 70}\n")

    for num, title, level in outline:
        indent = '  ' * (level - 1)
        if num:
            log_info(f'{indent}[{num}] {title}')
        else:
            log_info(f'{indent}[] {title}')  # 显式显示空编号

    from collections import Counter
    level_dist = Counter(level for _, _, level in outline)
    dist_str = "  ".join(f"L{lvl}:{level_dist[lvl]}" for lvl in sorted(level_dist.keys()))
    log_info(f"✓ 共 {len(outline)} 项 | 级别分布: {dist_str}")


if __name__ == "__main__":
    import sys

    doc_path = sys.argv[1] if len(sys.argv) > 1 else input("请输入Word文档路径: ")
    log_info('正在提取 Word 大纲（OutlineLevel 1-9）...')
    try:
        outline = extract_word_outline(doc_path, max_level=9)
        print_word_outline(outline)
    except Exception as e:
        log_error(f'❌ 提取失败: {str(e)}')
        import traceback

        traceback.print_exc()
        sys.exit(1)