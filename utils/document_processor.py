import os
import re
import time
import pandas as pd
from docx import Document
import tempfile
import shutil
import sys
from pathlib import Path
# 初评报告路径统一由 utils/initial_review_report.py 生成（在各校验函数内部按需导入）

# 确保可以导入 extend 目录下的模块
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.append(project_root)

from extend.hierarchical_matcher import HierarchicalMatcher
from extend.matcher_config import MatcherConfig
from utils.report_generator import ReportGenerator
from utils.runtime_logger import RuntimeLogger  # ✅ 新增导入
from datetime import datetime
import json
from utils.runtime_logger import log_debug, log_error, log_info, log_warn


class DocumentProcessor:
    """文档处理类，用于提取 Word 和 Excel 的内容结构"""

    # 类级缓存
    # _temp_files: [(路径, 属主项目名或None)]；属主来自线程局部项目上下文，
    #              并行任务只清理自己的临时文件，避免删掉其他运行中任务的文件
    _temp_files = []
    _cache = {}

    @staticmethod
    def _register_temp(path):
        """注册临时文件/目录，属主为当前线程的项目名（未设置则为 None）"""
        try:
            from utils.runtime_logger import RuntimeLogger
            owner = RuntimeLogger.get_project()
        except Exception:
            owner = None
        DocumentProcessor._temp_files.append((path, owner))

    @staticmethod
    def clear_cache(owner=None, max_age_hours=None):
        """
        清理临时文件和缓存。

        :param owner: 只清理指定属主（任务）注册的临时文件 —— 并行任务安全
        :param max_age_hours: 只清理早于该时长的陈旧临时文件 —— 跨次运行清理
        :return: 两者都未指定时执行全量清理（兼容旧行为；并行任务运行期间禁止调用）
        """
        import time as _time
        cutoff = _time.time() - max_age_hours * 3600 if max_age_hours else None
        remaining = []
        for entry in DocumentProcessor._temp_files:
            if isinstance(entry, tuple):
                temp_file, entry_owner = entry
            else:  # 兼容旧格式（纯路径）
                temp_file, entry_owner = entry, None
            if owner is not None and entry_owner != owner:
                remaining.append(entry)
                continue
            if cutoff is not None:
                try:
                    if os.path.getmtime(temp_file) >= cutoff:
                        # 陈旧度不满足：保留（可能正被其他任务使用）
                        remaining.append(entry)
                        continue
                except OSError:
                    remaining.append(entry)
                    continue
            try:
                if os.path.exists(temp_file):
                    if os.path.isfile(temp_file):
                        os.remove(temp_file)
                    elif os.path.isdir(temp_file):
                        shutil.rmtree(temp_file, ignore_errors=True)
            except Exception as e:
                # 文件可能正被其他任务/COM占用：保留登记，留待下次清理
                remaining.append(entry)
                try:
                    from utils.runtime_logger import RuntimeLogger
                    RuntimeLogger.log(f"[CLEANUP] 无法删除临时文件 {temp_file}: {e}", "DEBUG")
                except Exception:
                    pass

        DocumentProcessor._temp_files = remaining
        if owner is None and max_age_hours is None:
            DocumentProcessor._cache = {}

    @staticmethod
    def cleanup_stale_temp(max_age_hours=6):
        """清理历史运行遗留的陈旧临时文件（并行安全：只动超过时限的文件）"""
        DocumentProcessor.clear_cache(max_age_hours=max_age_hours)

    @staticmethod
    def _convert_doc_to_docx(doc_path):
        """将 .doc 转换为临时 .docx 文件 (增加重试机制与后台干扰屏蔽)"""
        import win32com.client as win32
        import pythoncom
        import time

        tmp_copy_path = None
        tmp_result_path = None
        word = None
        doc = None

        try:
            pythoncom.CoInitialize()

            # 使用 DispatchEx 启动独立的进程，减少互相干扰
            try:
                word = win32.DispatchEx("Word.Application")
            except:
                word = win32.Dispatch("Word.Application")

            word.Visible = False
            word.DisplayAlerts = 0

            # 禁用可能导致 Word  busy 的后台任务
            word.Options.ConfirmConversions = False
            word.Options.SaveNormalPrompt = False
            word.Options.UpdateLinksAtOpen = False
            word.Options.CheckSpellingAsYouType = False
            word.Options.CheckGrammarAsYouType = False

            # 物理隔离
            abs_doc_path = os.path.abspath(doc_path)
            temp_dir = tempfile.gettempdir()
            tmp_copy_path = os.path.join(
                temp_dir, f"doc2docx_{os.urandom(4).hex()}.doc"
            )
            shutil.copy2(abs_doc_path, tmp_copy_path)

            log_info(f'[CONVERT] 正在打开副本: {tmp_copy_path}')

            doc = word.Documents.Open(
                FileName=tmp_copy_path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )

            # 关键：给 Word 一些缓冲时间，防止 RPC_E_CALL_REJECTED
            time.sleep(1.0)

            # 准备输出路径
            tmp_result_path = tmp_copy_path + "x"

            # 带有重试机制的保存操作
            max_retries = 5
            last_err = None
            for i in range(max_retries):
                try:
                    log_info(f'[CONVERT] 正在执行转存 (尝试 {i + 1}/{max_retries}): {tmp_result_path}')
                    # FileFormat=16 是 wdFormatXMLDocument
                    doc.SaveAs2(tmp_result_path, FileFormat=16)
                    break
                except Exception as e:
                    last_err = e
                    # 如果错误码是 -2147418111 (Call rejected)，则进行重试
                    if "拒绝接收呼叫" in str(e) or "-2147418111" in str(e):
                        wait_time = 0.5 * (i + 1)
                        log_warn(f'[CONVERT] Word 忙碌，{wait_time}s 后重试...')

                        time.sleep(wait_time)
                        continue
                    else:
                        raise e
            else:
                # 循环结束未 break 表示全败
                raise last_err

            doc.Close(False)
            doc = None  # 显式释放

            return tmp_result_path

        except Exception as e:
            log_error(f'[CONVERT] ❌ 转换最终失败: {e}')
            import traceback

            traceback.print_exc()
            return None
        finally:
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            if word:
                try:
                    if word.Documents.Count == 0:
                        word.Quit()
                except:
                    pass

            # 清理
            if tmp_copy_path and os.path.exists(tmp_copy_path):
                try:
                    os.remove(tmp_copy_path)
                except:
                    pass

            pythoncom.CoUninitialize()

    @staticmethod
    def extract_word_outline_stable_win32(
        file_path, max_level=9, progress_callback=None
    ):
        """
        稳定的大纲提取方法 (基于 OutlineLevel)，完美解决编号丢失和正文过滤问题。
        对应 test_oval.py 的优化版本，支持提取正文内容。
        """
        import win32com.client as win32
        import pythoncom

        # 初始化计时
        start_time = time.time()
        logger = RuntimeLogger()
        logger.log(
            f"开始Win32方式提取文档结构: {os.path.basename(file_path)}", level="INFO"
        )

        pythoncom.CoInitialize()
        word = None
        doc = None
        sections = []

        try:
            abs_path = os.path.abspath(str(file_path))
            if not os.path.exists(abs_path):
                return []

            # 启动 Word
            try:
                word = win32.DispatchEx("Word.Application")
            except:
                word = win32.Dispatch("Word.Application")

            word.Visible = False
            word.DisplayAlerts = 0

            doc = word.Documents.Open(
                FileName=abs_path,
                ReadOnly=True,
                AddToRecentFiles=False,
                ConfirmConversions=False,
                Visible=False,
            )

            # 刷新大纲视图环境
            try:
                word.ActiveWindow.View.Type = 3  # wdOutlineView
            except:
                pass

            # 自动编号计数器 (支持 1-9 级大纲)
            auto_counters = [0] * 10

            current_section = {
                "level": 0,
                "title": "前言/未归类",
                "num": "",
                "content": [],
                "full_path": "前言",
            }

            # --- 优化：使用 Paragraphs 集合的 Count 属性预获取总数，并添加进度汇报 ---
            try:
                total_paras = doc.Paragraphs.Count
            except:
                total_paras = 0

            log_debug(f'[PROCESS] 文档预估段落数: {total_paras}，开始提取...')

            # 初始化进度计数器
            p_idx = 0
            content_lines_added = 0

            for para in doc.Paragraphs:
                try:
                    p_idx += 1
                    level = para.OutlineLevel
                    text = para.Range.Text

                    if not text or text.strip() in ["\r", "\x07", ""]:
                        continue

                    if p_idx % 500 == 0:
                        progress_msg = (
                            f"[PROCESS] 提取进度: {p_idx}/{total_paras}"
                            if total_paras > 0
                            else f"[PROCESS] 已提取 {p_idx} 段"
                        )
                        log_info(progress_msg)
                        if progress_callback:
                            percent = (
                                int((p_idx / total_paras) * 100)
                                if total_paras > 0
                                else 0
                            )
                            progress_callback(percent, progress_msg)

                    # 1-9 级大纲项
                    if (1 <= level <= 9 and level <= max_level) or (
                        level >= 10 and para.Range.Font.Bold and len(text.strip()) < 50
                    ):
                        # 如果是正文但加粗且短，视为 4 级标题（或者比当前章节深一级）
                        is_pseudo_heading = False
                        if level >= 10:
                            is_pseudo_heading = True
                            level = 4  # 默认设为4级
                            if current_section and current_section["level"] >= 1:
                                level = min(9, current_section["level"] + 1)

                        # 1. 更新结构化计数器
                        auto_counters[level] += 1
                        for i in range(level + 1, 10):
                            auto_counters[i] = 0
                        # 生成默认结构编号 (如 1.2.1)
                        struct_num = ".".join(
                            str(auto_counters[i]) for i in range(1, level + 1)
                        )

                        # 保存上一个章节
                        if (
                            current_section["title"] != "前言/未归类"
                            or current_section["content"]
                        ):
                            current_section["content"] = "\n".join(
                                current_section["content"]
                            ).strip()
                            sections.append(current_section)

                        # 处理新章节标题
                        try:
                            list_text = str(para.Range.ListFormat.ListString).strip()
                        except:
                            list_text = ""

                        title = text.strip().replace("\r", "").replace("\x07", "")

                        num_clean = ""
                        if list_text:
                            num_clean = (
                                list_text.rstrip(".")
                                .rstrip(")")
                                .rstrip("：")
                                .rstrip(":")
                                .strip()
                            )
                            if title.startswith(list_text):
                                title = title[len(list_text) :].strip()
                            elif num_clean and title.startswith(num_clean):
                                title = re.sub(
                                    rf"^{re.escape(num_clean)}[\s\.．、]*", "", title
                                ).strip()
                        else:
                            # 尝试从标题提取手动编号
                            num_match = re.match(r"^([\d\.]+)\s*(.*)", title)
                            if num_match:
                                candidate = num_match.group(1).rstrip(".")
                                # 验证改进：含有点的多级编号，或短小的单级编号
                                if "." in candidate or len(candidate) <= 2:
                                    num_clean = candidate
                                    title = num_match.group(2).strip()
                            else:
                                # 尝试匹配中文编号，如 "一、", "第一章", "1.1" (全角)
                                cn_num_match = re.match(
                                    r"^([第]?[一二三四五六七八九十百]+[章节]?|[0-9\.]+)[、\.\s]",
                                    title,
                                )
                                if cn_num_match:
                                    num_clean = cn_num_match.group(1).strip()
                                    title = title[cn_num_match.end() :].strip()

                        # 2. 关键优化：如果未提取到显式编号，使用结构化编号补全
                        if not num_clean:
                            num_clean = struct_num
                        else:
                            # 核心修正：如果解析出的编号层级与大纲级别（Style）不符，以编号为准（处理乱序文档）
                            # 例如：4.1.1 看起来是 3 级，但如果样式被设为了 Heading 1，则 level 会是 1
                            inferred_level = num_clean.count(".") + 1
                            if inferred_level != level and inferred_level <= 9:
                                level = inferred_level

                            # 如果提取到了显式编号，尝试同步计数器，保证后续自动编号的连续性
                            try:
                                parts = [
                                    int(p) for p in num_clean.split(".") if p.isdigit()
                                ]
                                if len(parts) == level:
                                    for i, p_val in enumerate(parts):
                                        auto_counters[i + 1] = p_val
                            except:
                                pass

                        # 关键：title 只保留核心文本以提高匹配成功率，num 专门存储编号
                        final_title = f"{num_clean} {title}" if num_clean else title
                        current_section = {
                            "level": level,
                            "title": title,  # 仅标题文本
                            "num": num_clean,
                            "display": final_title,  # 完整显示文本
                            "content": [],
                            "full_path": "",  # 后续统一计算
                        }
                    else:
                        # 正文内容
                        content_text = (
                            text.strip().replace("\r", "").replace("\x07", "")
                        )
                        if content_text:
                            current_section["content"].append(content_text)
                            content_lines_added += 1

                            # 每500行内容输出一次进度
                            if content_lines_added % 500 == 0:
                                elapsed = time.time() - start_time
                                logger.log(
                                    f"内容提取进度 - 已处理段落: {p_idx}, 添加内容行: {content_lines_added}, 耗时: {elapsed:.1f}秒",
                                    level="INFO",
                                )
                except:
                    continue

            # 保存最后一个章节
            if current_section["title"] != "前言/未归类" or current_section["content"]:
                current_section["content"] = "\n".join(
                    current_section["content"]
                ).strip()
                sections.append(current_section)

            # 补充 full_path
            current_titles = {}
            for s in sections:
                lvl = s["level"]
                current_titles[lvl] = s["title"]
                for i in range(lvl + 1, 10):
                    current_titles[i] = ""
                path_parts = [
                    current_titles[i]
                    for i in range(1, lvl + 1)
                    if i in current_titles and current_titles[i]
                ]
                s["full_path"] = " > ".join(path_parts)

            # 输出详细的处理统计
            total_elapsed = time.time() - start_time
            logger.log(
                f"第一轮提取完成 - 总段落: {p_idx}, 章节数: {len(sections)}, 内容行数: {content_lines_added}, 总耗时: {total_elapsed:.2f}秒",
                level="INFO",
            )

            log_info(f'[STABLE-EXTRACT] ✓ 成功提取 {len(sections)} 个章节')
            return sections

        except Exception as e:
            log_error(f'[STABLE-EXTRACT] ❌ 失败: {e}')
            import traceback

            traceback.print_exc()
            return []
        finally:
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            if word:
                try:
                    if word.Documents.Count == 0:
                        word.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    @staticmethod
    def load_word_document(file_path):
        """加载 Word 文档并返回 Document 对象"""
        file_path_str = str(file_path)
        try:
            if file_path_str.lower().endswith(".doc"):
                temp_docx = DocumentProcessor._convert_doc_to_docx(file_path_str)
                if temp_docx:
                    # 记录临时文件供后续清理
                    DocumentProcessor._register_temp(temp_docx)
                    return Document(temp_docx)
                else:
                    raise ValueError(f"Failed to convert .doc file: {file_path_str}")
            else:
                return Document(file_path_str)
        except Exception as e:
            log_error(f'[ERROR] 加载 Word 文档失败: {file_path_str} - {e}')
            raise

    @staticmethod
    def load_excel_workbook(file_path, cache=True):
        """加载 Excel 工作簿（cache 参数兼容保留；工作簿缓存只写不读且大文件下
        纯属内存泄漏，已停用存储，每次按需重新加载）"""
        file_path_str = str(file_path)
        try:
            # 使用 openpyxl 或 pandas 加载 Excel
            import openpyxl

            wb = openpyxl.load_workbook(file_path_str, data_only=False)
            return wb
        except Exception as e:
            log_error(f'[ERROR] 加载 Excel 工作簿失败: {file_path_str} - {e}')
            raise

    @staticmethod
    def extract_images_from_word(file_path):
        """
        从 Word 文档中提取图片。
        策略：先尝试寻找“功能架构图”相关章节提取图片，如果失败则提取全部图片。
        """
        import os
        import tempfile
        import zipfile
        from docx import Document
        from docx.oxml.ns import qn

        temp_dir = tempfile.mkdtemp()
        image_paths = []
        
        try:
            # 1. 尝试定位目标章节
            doc = Document(file_path)
            keywords = ["功能架构图", "业务架构", "逻辑架构", "系统架构", "架构图", "功能结构"]
            
            target_rids = []
            found_section = False
            
            for para in doc.paragraphs:
                text = para.text.strip()
                # 匹配章节标题
                if any(kw in text for kw in keywords) and (para.style.name.startswith('Heading') or para.style.name == 'Title'):
                    found_section = True
                    continue
                
                # 如果已在章节内，遇到下一个标题则可能结束（简化处理：持续提取直到遇到图片或下一个大章节）
                if found_section:
                    # 查找当前段落的所有图片引用
                    blips = para._element.xpath('.//a:blip')
                    for blip in blips:
                        rid = blip.get(qn('r:embed'))
                        if rid: target_rids.append(rid)
                    
                    # 如果进入了明显不相关的章节标题，可以退出，但为了保险我们这里多找几段
            
            # 提取选定的 rId
            for rid in target_rids:
                try:
                    part = doc.part.related_parts[rid]
                    ext = part.partname.split('.')[-1]
                    img_name = f"section_{len(image_paths)+1}.{ext}"
                    img_path = os.path.join(temp_dir, img_name)
                    with open(img_path, 'wb') as f:
                        f.write(part.blob)
                    image_paths.append(img_path)
                except:
                    continue
        except Exception as e:
            log_error(f'[EXTRACT-DOCX] 章节定位逻辑异常: {e}')

        # 2. 兜底：如果章节内没找到，或者文档不是 .docx，则通过 Zip 提取所有媒体
        if not image_paths:
            try:
                with zipfile.ZipFile(file_path, 'r') as zf:
                    for item in zf.namelist():
                        if item.startswith('word/media/'):
                            ext = item.split('.')[-1].lower()
                            if ext in ['png', 'jpg', 'jpeg', 'bmp', 'gif']:
                                img_name = os.path.basename(item)
                                img_path = os.path.join(temp_dir, img_name)
                                with open(img_path, 'wb') as f:
                                    f.write(zf.read(item))
                                image_paths.append(img_path)
            except Exception as e:
                log_error(f'[EXTRACT-DOCX] ZIP全量提取失败: {e}')

        # 记录临时目录以便后续清理
        DocumentProcessor._register_temp(temp_dir)
        # 去重返回
        return sorted(list(set(image_paths)))

    @staticmethod
    def extract_word_structure(file_path, use_stable=False, progress_callback=None):
        """
        提取 Word 文档的全层级标题及正文
        支持文件路径（str）或已加载的 Document 对象
        """
        import re  # 确保在函数作用域内可以访问re模块
        import time

        start_time = time.time()

        try:
            from utils.runtime_logger import RuntimeLogger

            if RuntimeLogger:
                RuntimeLogger.log(f"📜 开始提取Word文档结构...", level="INFO")
        except:
            pass

        # 提取文件名用于日志
        if hasattr(file_path, "paragraphs"):
            doc_name = "内存文档对象"
        else:
            doc_name = os.path.basename(str(file_path))

        log_info(f'[DOC-PROCESS-START] 开始处理文档: {doc_name}')

        # 如果启用稳定大纲提取模式 (基于 Word COM)
        if use_stable and isinstance(file_path, (str, Path)):
            log_info(f'[DOC-PROCESS] 启用稳定大纲提取模式: {file_path}')
            stable_start = time.time()

            try:
                result = DocumentProcessor.extract_word_outline_stable_win32(
                    file_path, progress_callback=progress_callback
                )
                stable_duration = time.time() - stable_start

                if result and len(result) > 0:
                    total_duration = time.time() - start_time
                    log_info(f'[DOC-PROCESS-SUCCESS] 稳定大纲提取成功: {len(result)}项, 耗时: {stable_duration:.2f}秒, 总耗时: {total_duration:.2f}秒')
                    return result
                else:
                    log_info(f'[DOC-PROCESS-FALLBACK] 稳定大纲提取无结果，回退到传统方法')
            except Exception as e:
                stable_duration = time.time() - stable_start
                log_error(f'[DOC-PROCESS-ERROR] 稳定大纲提取失败(耗时{stable_duration:.2f}秒): {e}')
                log_error(f'[DOC-PROCESS-FALLBACK] 回退到传统方法')

        log_info(f'[DOC-PROCESS] 开始直接大纲级别分析...')

        temp_docx = None
        doc = None

        # 检查是否已是 Document 对象（通过类名检查，避免 isinstance 作用域问题）
        doc_open_start = time.time()
        if hasattr(file_path, "paragraphs") and hasattr(file_path, "element"):
            # 这是一个 Document 对象
            doc = file_path
            log_info(f'[DOC-PROCESS] 使用传入的Document对象，段落数: {len(doc.paragraphs)}')
        else:
            # 统一转为字符串处理，防止 pathlib.Path 对象导致 lower() 失败
            file_path_str = str(file_path)
            log_info(f'[DOC-PROCESS] 正在打开文件: {file_path_str}')
            if file_path_str.lower().endswith(".doc"):
                log_info(f'[DOC-PROCESS] 检测到.doc格式，启动COM转换...')
                convert_start = time.time()
                temp_docx = DocumentProcessor._convert_doc_to_docx(file_path_str)
                convert_duration = time.time() - convert_start
                if not temp_docx:
                    log_error(f'[DOC-PROCESS] ❌ .doc转换失败(耗时{convert_duration:.2f}秒)')
                    return []
                log_info(f'[DOC-PROCESS] .doc转换成功(耗时{convert_duration:.2f}秒)，临时文件: {temp_docx}')
                doc_to_read = temp_docx
            else:
                doc_to_read = file_path_str

            log_info(f'[DOC-PROCESS] 正在加载docx对象: {doc_to_read}')
            load_start = time.time()
            doc = Document(doc_to_read)
            load_duration = time.time() - load_start
            log_info(f'[DOC-PROCESS] docx加载完成(耗时{load_duration:.2f}秒)，段落数: {len(doc.paragraphs)}')

        doc_open_duration = time.time() - doc_open_start
        log_info(f'[DOC-PROCESS] 文档打开总耗时: {doc_open_duration:.2f}秒')

        try:
            sections = []

            # 预识别文档是否使用了“标题”样式簇
            has_heading_styles = False
            for p in doc.paragraphs[:300]:
                sn = p.style.name.lower()
                if "heading" in sn or "标题" in sn:
                    if "toc" not in sn and "目录" not in sn:
                        has_heading_styles = True
                        break

            # 扩展计数器，防止深层标题越界
            level_counters = [0] * 10
            current_titles = {i: "" for i in range(1, 10)}

            # 新增：层级计数器 (用于模拟 Word 自动编号)
            current_section = {
                "level": 0,
                "title": "前言/未归类",
                "num": "",
                "content": [],
                "full_path": "前言",
            }

            # 【新增】用于跟踪已出现的章节,避免重复(基于编号+标题的完全匹配)
            seen_chapters = set()  # 存储 (编号, 标题) 的元组
            used_numbers = set()  # 存储已使用的编号（用于检测编号冲突）

            toc_styles = ["toc", "目录", "TOC"]

            # 常见核心章节关键字
            core_keywords = [
                "需求说明",
                "系统现状",
                "概况",
                "已实现功能",
                "存在问题",
                "总体描述",
                "建设目标",
                "建设必要性",
                "功能架构图",
                "功能需求",
                "关键时序图",
                "业务逻辑图",
                "功能描述",
                "需求功能清单",
                "项目需求清单",
                "附加值",
                "调整因子",
                "规模因子",
                "应用类型",
                "质量及特性",
                "开发语言",
                "团队背景",
                "完整性",
            ]

            # 引导性文字黑名单 (防止误判为标题)
            instruction_blacklist = [
                "请在此处",
                "说明本项目",
                "示例（",
                "示例:",
                "图3.",
                "图4.",
                "本工程应",
                "建议进行",
            ]

            # 【优化】不再重新映射Heading层级，直接使用Heading的原始数字
            # 初始化所有可能的 Heading 1-9 映射，避免在大文档中全量扫描样式
            heading_styles_map = {i: i for i in range(1, 10)}
            log_info(f'[HEADING-MAP] 应用默认Heading样式映射: {heading_styles_map}')

            def get_level_enhanced(text, style_name, p_obj):
                if any(kw in text for kw in instruction_blacklist):
                    return None

                style_lower = style_name.lower()
                is_heading_style = "heading" in style_lower or "标题" in style_lower

                # 核心过滤：非标题样式通常不以末尾标点结束
                # 如果是标题样式，则放宽标点限制
                t_strip = text.strip()
                if not is_heading_style:
                    if t_strip.endswith(
                        ("；", ";", "。", "，", ",", "！", "!", "：", ":", "？", "?")
                    ):
                        return None

                if len(t_strip) > 100:  # 放宽到 100 字
                    return None

                style_lower = style_name.lower()
                is_heading_style = "heading" in style_lower or "标题" in style_lower

                # 获取段落属性对象
                pPr = None
                try:
                    pPr = p_obj._element.pPr
                except:
                    pass

                # 提取手动编号 (提前提取，用于覆盖大纲级别)
                # 增强正则：支持 [4.1.1.1] 或 (4.1.1.1) 这种带括号的编号
                manual_num_match = re.match(
                    r"^\s*[\[\(]?(\d+([\.\．]\d+)+)[\]\)]?[\.．\s]*", text
                )
                manual_level = None
                if manual_num_match:
                    clean_num = manual_num_match.group(1).replace("．", ".")
                    manual_level = len(clean_num.split("."))

                # 提前计算各种特征属性
                is_bold = False
                try:
                    for run in p_obj.runs:
                        if run.text.strip() and run.bold:
                            is_bold = True
                            break
                    if (
                        not is_bold
                        and hasattr(p_obj.style, "font")
                        and p_obj.style.font
                        and p_obj.style.font.bold
                    ):
                        is_bold = True
                except:
                    pass

                is_core = any(kw == text.strip().rstrip(":：") for kw in core_keywords)
                # 标题通常不包含“本段”、“如下图”等引导词
                if any(kw in text for kw in ["本段", "如下图", "下表", "如下所示"]):
                    is_core = False

                is_very_short = len(text.strip()) < 35

                # --- 决策逻辑 ---

                # 0. 增加：如果文本内容仅包含“关键时序图”、“功能描述”等正文关键字，且没加粗，坚决不作为标题
                # 注意：即便带了手动编号（如 4.1.1.1 关键时序图），也要在步骤1中精准拦截
                if text.strip().rstrip(":：") in [
                    "关键时序图",
                    "业务逻辑图",
                    "关键时序图/业务逻辑图",
                    "功能描述",
                ]:
                    if not is_bold and not is_heading_style:
                        return None

                # 1. 【优先】检查大纲级别 (Word XML outline level)
                # 完全模拟WPS大纲视图：只要段落有大纲级别，就应该被识别为标题
                # WPS智能目录/大纲视图就是基于这个属性，不依赖样式、加粗等其他特征
                # 【重要】outline level是Word对文档结构的官方定义，应该绝对信任
                if pPr is not None:
                    try:
                        if pPr.outlineLvl is not None:
                            val = int(pPr.outlineLvl.val)

                            # 【边界检查】过滤异常的深层级（>8通常是数据错误）
                            if val >= 9:
                                return None

                            # 【关键修复】完全信任outline level
                            # Word的outline level: 0-8对应9个层级
                            # 我们映射为: 0→1, 1→2, ..., 8→9
                            outline_level = val + 1

                            # 【只有一个严格的过滤条件】：
                            # 如果文本超长（>150字符）且是normal样式，可能是误标的正文
                            # 但我们优先相信outline level，只在明显异常时才过滤
                            if len(text.strip()) > 150 and style_lower == "normal":
                                # 仍然返回outline level，而不是None
                                # 让后续逻辑来判断是否需要处理
                                pass

                            # 其他所有有大纲级别的段落都识别为标题（模拟WPS大纲视图）
                            return outline_level
                    except:
                        pass

                # 2. 如果有明确的多级手动编号，使用它（但仅在没有outline level时）
                # 增强正则：支持 1.1, 1.1.1, 1 . 1, 1.1.1.1 等，允许括号、空格和结尾点
                # 【修复】编号后可以直接跟汉字，如"4.1.1.1新增业务工单"
                manual_num_match = re.match(
                    r"^\s*[\[\(]?(\d+([\.\．\s]*[\.．]\s*\d+)+)[\]\)]?([\.．\s]+|(?=[\u4e00-\u9fa5]))",
                    text.strip(),
                )
                if manual_num_match:
                    num_str = (
                        manual_num_match.group(1).replace(" ", "").replace("．", ".")
                    )
                    parts = [p for p in num_str.split(".") if p.strip()]
                    if parts:
                        num_level = len(parts)
                        # --- 拦截 4.1.1.1 这种带编号的正文固定小标题 ---
                        if num_level >= 4:
                            remain = (
                                text.strip()[len(manual_num_match.group(0)) :]
                                .strip()
                                .rstrip(":：")
                            )
                            if remain in [
                                "关键时序图",
                                "业务逻辑图",
                                "关键时序图/业务逻辑图",
                                "功能描述",
                                "功能说明",
                            ]:
                                if not is_bold and not is_heading_style:
                                    return None
                        return num_level

                # 3. 自动编号判定
                if pPr is not None:
                    try:
                        if pPr.numPr is not None:
                            if any(
                                s in style_lower
                                for s in ["list", "列表", "bullet", "indent", "缩进"]
                            ):
                                return None
                            ilvl_obj = pPr.numPr.ilvl
                            if ilvl_obj is not None:
                                at_level = int(ilvl_obj.val) + 1
                                if at_level == 1:
                                    if not (
                                        is_heading_style
                                        or is_bold
                                        or is_core
                                        or is_very_short
                                    ):
                                        return None
                                return at_level
                    except:
                        pass

                # 4. 样式判定（Heading 1-9）
                if is_heading_style:
                    detected_level = None
                    for i in range(1, 10):
                        if str(i) in style_lower:
                            detected_level = i
                            break

                    # 特殊拦截：Heading 6 且内容是"关键时序图"等正文小标题，不予识别
                    if detected_level == 6:
                        pure_text = text.strip().rstrip(":：")
                        if pure_text in [
                            "关键时序图",
                            "业务逻辑图",
                            "关键时序图/业务逻辑图",
                            "功能描述",
                            "功能说明",
                        ]:
                            return None

                    if detected_level:
                        # 【关键修复】：使用映射表将Heading样式编号转换为实际层级
                        # 例如：如果文档只使用Heading 2,3,4,5,6
                        # 那么 Heading 2 → Level 1, Heading 3 → Level 2, 等等
                        if detected_level in heading_styles_map:
                            return heading_styles_map[detected_level]
                        else:
                            # 如果不在映射表中（理论上不应该发生），使用原值
                            return detected_level

                # 5. 单级手动编号 (1. 或 一、)
                if re.match(r"^(\d+|[一二三四五六七八九十]+)[\.．、]\s*", text):
                    # 【放宽限制】模拟WPS智能目录：只要有编号格式，长度<60，不以标点结尾，就识别为标题
                    if len(text) > 60:  # 放宽到60字符
                        return None

                    # 如果是加粗或标题样式，直接识别
                    if is_heading_style or is_bold:
                        return 1

                    # 【重要修复】如果文档中有标题样式，则简单编号必须也使用标题样式或加粗
                    if has_heading_styles and not is_heading_style and not is_bold:
                        # 但如果是核心关键词或很短的文本，仍然识别
                        if not (is_core or is_very_short):
                            return None

                    if not (
                        is_heading_style
                        or is_bold
                        or is_core
                        or (is_very_short and "list" not in style_lower)
                    ):
                        return None
                    return 1

                # 6. 无点编号: "1 需求说明"
                if re.match(r"^(\d+|[一二三四五六七八九十]+)\s+[\u4e00-\u9fa5]", text):
                    # 如果有 Heading 样式，则这种非标准的无点编号必须带有 Heading 样式，否则视为普通正文
                    if has_heading_styles and not is_heading_style:
                        return None
                    if is_bold or (is_core and is_very_short):
                        return 1

                # 7. 启发式保底 (只有在确定没有标准标题样式时才放宽)
                if not has_heading_styles:
                    if is_core and len(text) < 30:
                        return 2
                    if len(text) < 50 and (is_bold or (is_core and len(text) < 15)):
                        return 2 if is_core else 3
                else:
                    # 如果有标准标题样式，则只有显式的标题样式或极其明显的核心章节才作为标题
                    # 严禁将 Normal 样式的短句（如表格标题、表格内容）识别为层级
                    if style_lower == "normal" and not is_bold:
                        return None

                    if (
                        (is_core or is_bold)
                        and len(text) < 35
                        and "：" not in text
                        and ":" not in text
                    ):
                        # 仅在非列表样式下考虑
                        if "list" not in style_lower and "normal" not in style_lower:
                            return 4

                return None

            # 跟踪已提取的标题，避免重复
            # TOC编号映射表：标题文本 -> (编号, 层级)
            toc_number_map = {}

            # 【优化】延迟检测 outline level，在后续遍历中自动处理
            uses_outline_level = False

            # 收集所有段落（使用缓存以提升性能）
            all_paragraphs = []
            # 预采样检测 outline level
            for p in doc.paragraphs[:300]:
                try:
                    if (
                        p._element.pPr is not None
                        and p._element.pPr.outlineLvl is not None
                    ):
                        uses_outline_level = True
                        break
                except:
                    pass

            if uses_outline_level:
                log_info('[INFO] 采样检测到文档使用 outline level')

            # 将表格内的段落也收集起来
            for p in doc.paragraphs:
                all_paragraphs.append(p)

            def collect_from_table(table):
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            all_paragraphs.append(p)
                        for nested in cell.tables:
                            collect_from_table(nested)

            for t in doc.tables:
                collect_from_table(t)

            # --- 核心性能优化：预提取段落属性 ---
            log_info(f'[DOC-PROCESS] 正在扫描 {len(all_paragraphs)} 个段落属性...')
            para_attr_cache = []
            style_id_to_name = (
                {}
            )  # 性能优化：缓存样式 ID 到名称的映射，避免昂贵的 p.style 访问

            for p in all_paragraphs:
                try:
                    # 1. 提取文本
                    t = p.text

                    # 2. 性能优化：通过 style_id 进行样式名称缓存
                    # p.style_id 是属性访问，极快；p.style 是复杂对象查找，极慢
                    sid = p.style_id
                    if sid not in style_id_to_name:
                        try:
                            style_id_to_name[sid] = p.style.name
                        except:
                            style_id_to_name[sid] = "Normal"
                    s = style_id_to_name[sid]

                    para_attr_cache.append(
                        {
                            "obj": p,
                            "text": t,
                            "text_strip": t.strip(),
                            "style": s,
                            "style_lower": s.lower(),
                        }
                    )
                except:
                    # 某些损坏的段落可能报错
                    para_attr_cache.append(
                        {
                            "obj": p,
                            "text": "",
                            "text_strip": "",
                            "style": "Normal",
                            "style_lower": "normal",
                        }
                    )

            # 第一轮：扫描所有TOC样式，建立映射表
            toc_entries_list = []
            for p_data in para_attr_cache:
                text = p_data["text_strip"]
                if not text:
                    continue

                style_name = p_data["style"]
                style_lower = p_data["style_lower"]

                # 过滤目录页码行
                if re.search(r"\.{3,}\s*\d+\s*$", text):
                    continue

                # 检查是否是TOC样式
                is_toc_style = any(s in style_lower for s in toc_styles)
                if is_toc_style:
                    # 提取toc样式的编号和标题
                    toc_match = re.match(
                        r"^(\d+(?:\.\d+)*)[.\s\t]+([\u4e00-\u9fa5\w\s()（）]+?)(?:\s+\d+)?$",
                        text,
                    )
                    if toc_match:
                        toc_number = toc_match.group(1)
                        toc_title = toc_match.group(2).strip()
                        toc_level = len(toc_number.split("."))
                        toc_entries_list.append(
                            (toc_title, toc_number, toc_level, True, False)
                        )
                        toc_number_map[toc_title] = (toc_number, toc_level, True)
                    else:
                        if "toc 1" in style_lower:
                            unnumbered_title = re.sub(r"\s+\d+$", "", text).strip()
                            if (
                                unnumbered_title
                                and len(unnumbered_title) < 50
                                and not re.match(r"^\d", unnumbered_title)
                            ):
                                toc_entries_list.append(
                                    (unnumbered_title, None, 1, False, False)
                                )
                                toc_number_map[unnumbered_title] = (None, 1, False)

            # 第二轮：处理所有段落，应用TOC映射
            main_processing_start = time.time()
            log_info(f'[DOC-PROCESS] 第二轮处理: 正在逐段落解析并应用TOC映射...')

            in_procedure_section = False
            procedure_parent_level = 0
            list_item_counter = 0

            sections_found = 0
            total_paragraphs = len(para_attr_cache)
            progress_interval = max(1, total_paragraphs // 50)

            for para_idx, p_data in enumerate(para_attr_cache):
                paragraph = p_data["obj"]
                # 进度报告
                if para_idx % progress_interval == 0:
                    if progress_callback:
                        progress = min(40 + (para_idx / total_paragraphs) * 40, 80)
                        progress_callback(
                            progress, f"正在处理段落: {para_idx}/{total_paragraphs}"
                        )

                text = p_data["text_strip"]
                if not text:
                    continue

                style_name = p_data["style"]
                style_lower = p_data["style_lower"]

                if re.search(r"\.{3,}\s*\d+\s*$", text):
                    continue

                if any(s in style_lower for s in toc_styles):
                    continue

                # 【增强】如果正文标题在TOC映射表中，使用TOC的编号
                # 支持多种匹配模式：
                # 1. 精确匹配：text == toc_title
                # 2. 去前缀匹配：text == "1 需求说明", toc_title == "需求说明"
                # 3. 部分匹配（仅限Heading样式）：text == "需求说明", toc_title 可能在"1. 需求说明"中
                # 【关键修复】按顺序从列表中查找未使用的TOC条目，支持重复标题
                # 【增强】改进匹配策略：优先考虑编号继承关系
                matched_toc = None
                matched_toc_index = None

                # 先判断是否是Heading样式（部分匹配需要这个条件）
                is_heading_style = (
                    "heading" in style_name.lower() or "标题" in style_name.lower()
                )

                # 【新增】提前检测正文文本中是否包含明确的多级编号
                # 例如 "4.1.1.1 集团360画像"，如果有明确编号，优先用编号来匹配TOC
                text_explicit_num = None
                text_explicit_num_match = re.match(r"^\s*(\d+(?:\.\d+)+)\s+(.+)$", text)
                if text_explicit_num_match:
                    text_explicit_num = text_explicit_num_match.group(1)
                    text_title_only = text_explicit_num_match.group(2).strip()

                    # 【优先策略】如果正文有明确的多级编号，首先在TOC中查找相同编号
                    for idx, (
                        toc_title,
                        toc_number,
                        toc_level,
                        has_number,
                        used,
                    ) in enumerate(toc_entries_list):
                        if used:
                            continue
                        # 精确匹配编号
                        if toc_number and toc_number == text_explicit_num:
                            matched_toc = (
                                toc_title,
                                (toc_number, toc_level, has_number),
                            )
                            matched_toc_index = idx
                            log_info(f"[TOC-MATCH-BY-NUMBER] 通过编号精确匹配: '{text}' → TOC '{toc_number}'")
                            break

                # 模式1: 精确匹配 - 在列表中按顺序查找第一个未使用的匹配项
                if not matched_toc:
                    for idx, (
                        toc_title,
                        toc_number,
                        toc_level,
                        has_number,
                        used,
                    ) in enumerate(toc_entries_list):
                        if used:
                            continue
                        if text == toc_title:
                            matched_toc = (
                                toc_title,
                                (toc_number, toc_level, has_number),
                            )
                            matched_toc_index = idx
                            log_info(f"[TOC-MATCH-EXACT] 精确标题匹配: '{text}' → TOC '{toc_number}'")
                            break

                if not matched_toc:
                    # 模式2: 去除正文标题开头的编号后匹配
                    # 例如：正文"1 需求说明" → 去掉"1 " → "需求说明" → 匹配TOC
                    text_without_prefix = re.sub(r"^[\d\.]+\s+", "", text)
                    if text_without_prefix != text:
                        for idx, (
                            toc_title,
                            toc_number,
                            toc_level,
                            has_number,
                            used,
                        ) in enumerate(toc_entries_list):
                            if used:
                                continue
                            if text_without_prefix == toc_title:
                                # 注意：这种情况下，原文本已经有编号（如"1 需求说明"），
                                # 只需要验证编号是否匹配TOC，不需要重新添加编号
                                # 提取原文本中的编号
                                original_num_match = re.match(r"^([\d\.]+)", text)
                                if original_num_match:
                                    original_num = original_num_match.group(1).rstrip(
                                        "."
                                    )
                                    if original_num == toc_number:
                                        # 原编号与TOC编号一致，直接使用原文本
                                        log_info(f"[TOC-MATCH-PREFIX-BY-NUM] 去除前缀后匹配且编号一致: '{text}' → TOC '{toc_number}'")
                                        matched_toc = (
                                            toc_title,
                                            (toc_number, toc_level, has_number),
                                        )
                                        matched_toc_index = idx
                                        break
                                    else:
                                        # 原编号与TOC编号不一致，打日志但继续
                                        log_info(f"[TOC-MISMATCH-NUM] 标题匹配但编号不同: 原'{original_num}' vs TOC'{toc_number}'")
                                        # 继续尝试其他TOC条目
                                        continue
                                else:
                                    # 原文本无编号，使用TOC编号
                                    matched_toc = (
                                        toc_title,
                                        (toc_number, toc_level, has_number),
                                    )
                                    matched_toc_index = idx
                                    log_info(f"[TOC-MATCH-PREFIX] 去除前缀后匹配: '{text}' → TOC '{toc_number}'")
                                    break

                    if not matched_toc and is_heading_style:
                        # 模式3: 部分匹配（仅对Heading样式）
                        # 例如：正文"需求说明" → 匹配TOC中包含"需求说明"的条目
                        # 但要求长度相近，避免"电子工单查询页面"匹配到"电子工单查询页面PDF查询渠道"
                        for idx, (
                            toc_title,
                            toc_number,
                            toc_level,
                            has_number,
                            used,
                        ) in enumerate(toc_entries_list):
                            if used:
                                continue
                            # 检查正文标题是否是TOC标题去掉编号后的内容
                            toc_without_prefix = re.sub(r"^[\d\.]+\s+", "", toc_title)
                            # 要求：正文标题与TOC标题去前缀后完全相同，或长度差不超过3个字符
                            if text == toc_without_prefix:
                                matched_toc = (
                                    toc_title,
                                    (toc_number, toc_level, has_number),
                                )
                                matched_toc_index = idx
                                log_info(f"[TOC-MATCH-PARTIAL] 部分匹配: '{text}' → TOC '{toc_title}'")
                                break
                            elif (
                                len(text) >= 3
                                and abs(len(text) - len(toc_without_prefix)) <= 3
                                and text in toc_without_prefix
                            ):
                                # 宽松匹配：正文是TOC的子串，且长度相近
                                matched_toc = (
                                    toc_title,
                                    (toc_number, toc_level, has_number),
                                )
                                matched_toc_index = idx
                                log_info(f"[TOC-MATCH-PARTIAL-LOOSE] 宽松匹配: '{text}' → TOC '{toc_title}'")
                                break

                if matched_toc:
                    # 标记TOC条目为已使用
                    if matched_toc_index is not None:
                        toc_entries_list[matched_toc_index] = (
                            toc_entries_list[matched_toc_index][0],  # toc_title
                            toc_entries_list[matched_toc_index][1],  # toc_number
                            toc_entries_list[matched_toc_index][2],  # toc_level
                            toc_entries_list[matched_toc_index][3],  # has_number
                            True,  # used = True
                        )

                    toc_title_used, (toc_number, toc_level, has_number) = matched_toc

                    # 去除编号前缀，获取纯标题文本用于关键词匹配
                    text_without_number = re.sub(r"^[\d\.]+\s*", "", text).strip()

                    # 检查是否是"过程说明"、"功能过程"等特殊章节
                    procedure_keywords = [
                        "过程说明",
                        "功能过程",
                        "过程描述",
                        "功能过程说明",
                    ]
                    if any(
                        keyword in text_without_number for keyword in procedure_keywords
                    ):
                        in_procedure_section = True
                        procedure_parent_level = toc_level
                        list_item_counter = 0
                        log_info(f'[PROCEDURE] 进入过程说明章节: {text_without_number}, Level={toc_level}')
                    elif in_procedure_section and toc_level <= procedure_parent_level:
                        # 退出过程说明章节（遇到同级或更高级章节）
                        in_procedure_section = False
                        list_item_counter = 0
                        log_info(f'[PROCEDURE] 退出过程说明章节')

                    # 【处理无编号章节】如"版本历史"
                    if not has_number or toc_number is None:
                        log_info(f'[TOC-MATCH] 正文标题匹配到无编号TOC: {text}')
                        level = toc_level
                        # 无编号章节，直接使用原标题，不添加编号
                        final_title = text
                        use_toc_number = False  # 标记为无编号，跳过后续的TOC编号处理

                        # 【保存无编号章节】
                        log_info(f'[DEBUG-WORD-STRUCTURE] Level: {level} | FinalTitle: {final_title}')

                        # 保存前一个章节
                        if (
                            current_section["title"] != "前言/未归类"
                            or current_section["content"]
                        ):
                            current_section["content"] = "\n".join(
                                current_section["content"]
                            ).strip()
                            sections.append(current_section)
                            log_info(f"[APPEND-SECTION #{len(sections)}] Level={current_section['level']}, Title={current_section['title'][:60]}")

                        # 更新路径跟踪
                        current_titles[level] = final_title
                        for l in range(level + 1, 10):
                            current_titles[l] = ""

                        path_parts = [
                            current_titles[l]
                            for l in range(1, level + 1)
                            if current_titles[l]
                        ]

                        # 无编号章节，num字段为空
                        current_section = {
                            "level": level,
                            "title": final_title,
                            "num": "",
                            "content": [],
                            "full_path": " > ".join(path_parts),
                        }

                        # 跳过后续处理，继续下一个段落
                        continue
                    else:
                        log_info(f'[TOC-MATCH] 正文标题匹配到TOC: {text} -> {toc_number}')

                        # 【检测编号冲突】如果TOC编号已被其他章节使用，说明TOC有误
                        # 但对于无编号的TOC条目（如"附录A"），跳过冲突检测
                        use_toc_number = True
                        if toc_number and toc_number in used_numbers:
                            log_warn(f'[WARN] TOC编号冲突: [{toc_number}] 已被使用，改用自动编号')
                            # 标记不使用TOC编号，但仍然保持是章节，使用TOC的level
                            use_toc_number = False
                            # 使用TOC的层级，但用自动编号
                            level = toc_level

                            # 【边界检查】
                            if level >= len(level_counters):
                                log_warn(f'[WARN] 层级 {level} 过深，跳过: {text[:40]}')
                                continue

                            # 增加当前层级计数器
                            level_counters[level] += 1

                            # 重置所有更深层级的计数器
                            for deeper in range(level + 1, len(level_counters)):
                                level_counters[deeper] = 0

                            # 构建自动编号（跳过值为0的层级）
                            active_parts = [
                                level_counters[i]
                                for i in range(1, level + 1)
                                if level_counters[i] > 0
                            ]
                            auto_num = ".".join(map(str, active_parts))

                            # 清理标题文本
                            clean_text = re.sub(r"^[\d\.]+\s*", "", text).strip()
                            final_title = f"{auto_num} {clean_text}"

                            # 记录已使用的编号
                            used_numbers.add(auto_num)

                            log_info(f'[AUTO-NUM-FALLBACK] TOC冲突，使用自动编号: {auto_num}, Level: {level}')
                            log_info(f'    → Generated: {auto_num}, Counters[1-6]: {level_counters[1:7]}')
                            log_info(f'[DEBUG-WORD-STRUCTURE] Level: {level} | FinalTitle: {final_title}')

                            # 保存前一个章节
                            if (
                                current_section["title"] != "前言/未归类"
                                or current_section["content"]
                            ):
                                current_section["content"] = "\n".join(
                                    current_section["content"]
                                ).strip()
                                sections.append(current_section)
                                log_info(f"[APPEND-SECTION #{len(sections)}] Level={current_section['level']}, Title={current_section['title'][:60]}")

                            # 更新路径跟踪
                            current_titles[level] = final_title
                            for l in range(level + 1, 10):
                                current_titles[l] = ""

                            path_parts = [
                                current_titles[l]
                                for l in range(1, level + 1)
                                if current_titles[l]
                            ]

                            # 提取编号
                            num_match = re.match(r"^([\d\.]+)", final_title)
                            section_num = (
                                num_match.group(1).rstrip(".") if num_match else ""
                            )

                            current_section = {
                                "level": level,
                                "title": final_title,
                                "num": section_num,
                                "content": [],
                                "full_path": " > ".join(path_parts),
                            }

                            # 跳过后续处理，继续下一个段落
                            continue
                        else:
                            # 【防止重复内容】基于(编号,标题)完全匹配检测重复
                            # 只有编号和标题都相同时才算重复,如:
                            # "[4] 功能需求" vs "[4] 功能需求" → 重复 ✓
                            # "[2] 系统现状" vs "[2] 功能架构图" → 不重复 ✓
                            clean_text = re.sub(r"^[\d\.]+\s*", "", text).strip()
                            chapter_key = (toc_number, clean_text)

                            if chapter_key in seen_chapters:
                                log_warn(f'[WARN] 检测到重复章节(编号+标题完全相同)，跳过: [{toc_number}] {clean_text}')
                                continue

                            # 记录这个章节
                            seen_chapters.add(chapter_key)
                            # 【修复】只在有编号时才记录到used_numbers
                            if toc_number:
                                used_numbers.add(toc_number)  # 记录已使用的编号

                    # 如果TOC编号没有冲突，使用TOC编号
                    if matched_toc and use_toc_number:
                        # 【原有的向后跳转检测已移除】

                        # 使用TOC编号和层级，重写text（保留原始标题文本，加上TOC编号）
                        # 如果原文本已有编号，先去掉
                        clean_text = re.sub(r"^[\d\.]+\s+", "", text)

                        # 直接构造最终标题，不走后续手动编号流程
                        level = toc_level

                        # 【修复】处理无编号的TOC条目（如"附录A"）
                        if toc_number is None or toc_number == "":
                            # 没有编号，直接使用原文本
                            final_title = text
                        else:
                            # 保留原文档编号格式（原文本可能是"1需求说明"或"1. 需求说明"）
                            # 提取原始编号后的分隔符（空格、点号等）
                            sep_match = re.match(
                                rf"^{re.escape(toc_number)}([^\u4e00-\u9fa5]*)", text
                            )
                            if sep_match:
                                separator = sep_match.group(
                                    1
                                )  # 可能是"."或" "或".  "等
                                final_title = f"{toc_number}{separator}{clean_text}"
                            else:
                                # 如果原文没有编号，则使用TOC编号（这种情况下原文可能只是标题）
                                final_title = f"{toc_number} {clean_text}"

                            # 同步计数器（仅当有编号时）
                            parts = toc_number.split(".")
                            for i, p in enumerate(parts):
                                idx = i + 1
                                if idx < len(level_counters):
                                    try:
                                        level_counters[idx] = int(p)
                                    except:
                                        pass
                            # 重置子级
                            for i in range(len(parts) + 1, len(level_counters)):
                                level_counters[i] = 0

                        log_info(f'    → Using TOC numbering, Updated Counters[1-6]: {level_counters[1:7]}')
                        log_info(f'[DEBUG-WORD-STRUCTURE] Level: {level} | FinalTitle: {final_title}')

                        # 保存前一个章节
                        if (
                            current_section["title"] != "前言/未归类"
                            or current_section["content"]
                        ):
                            current_section["content"] = "\n".join(
                                current_section["content"]
                            ).strip()
                            sections.append(current_section)
                            log_info(f"[APPEND-SECTION #{len(sections)}] Level={current_section['level']}, Title={current_section['title'][:60]}")

                        # 更新路径跟踪
                        current_titles[level] = final_title
                        for l in range(level + 1, 10):
                            current_titles[l] = ""

                        path_parts = [
                            current_titles[l]
                            for l in range(1, level + 1)
                            if current_titles[l]
                        ]

                        # 提取编号(支持数字后的空格)
                        num_match = re.match(r"^([\d\.]+)", final_title)
                        section_num = (
                            num_match.group(1).rstrip(".") if num_match else ""
                        )

                        current_section = {
                            "level": level,
                            "title": final_title,
                            "num": section_num,
                            "content": [],
                            "full_path": " > ".join(path_parts),
                        }

                        # 跳过后续处理，继续下一个段落
                        continue
                else:
                    # 没有匹配到TOC，使用原有逻辑识别层级
                    level = get_level_enhanced(text, style_name, paragraph)

                    # 【关键修复】当没有匹配到TOC时，尝试从前面已提取的编号推导
                    # 如果当前项目的标题出现在某个已识别的TOC条目中，使用TOC的编号
                    if not matched_toc and level is not None:
                        # 尝试模糊匹配：查找包含当前文本的TOC条目
                        text_clean = (
                            re.sub(r"^[\d\.]+\s+", "", text).strip().rstrip(":：")
                        )
                        for (
                            toc_title,
                            toc_number,
                            toc_level,
                            has_number,
                            used,
                        ) in toc_entries_list:
                            if not used and text_clean and toc_title:
                                # 检查是否包含关系或相似
                                if (
                                    text_clean in toc_title
                                    or toc_title in text_clean
                                    or text_clean == toc_title
                                ):
                                    # 找到可能的匹配
                                    matched_toc = (
                                        toc_title,
                                        (toc_number, toc_level, has_number),
                                    )
                                    level = toc_level
                                    log_info(f"[TOC-FUZZY-MATCH] 模糊匹配找到TOC: '{text_clean}' → '{toc_number}'")
                                    break

                # 【列表项强制识别】
                # 在"过程说明"等章节下，强制识别"1. xxx"、"2. xxx"格式为列表项
                # 必须在level判定后、但在跳变约束前处理，避免被误判为其他层级
                if in_procedure_section and level is not None:
                    list_item_match = re.match(r"^(\d+)[.．]\s+(.+)$", text)
                    if list_item_match:
                        item_num = int(list_item_match.group(1))
                        # 检查是否是连续的列表项（1,2,3...）或重新开始的列表（1）
                        if item_num == 1 or item_num == list_item_counter + 1:
                            # 这是列表项，强制设置level和标记
                            list_item_counter = item_num
                            level = procedure_parent_level + 1
                            # 标记为列表项，后续不再经过level_counters自动编号
                            is_list_item = True
                            log_info(f'[LIST-ITEM] 强制识别列表项: {text[:40]} -> Level {level}, Item#{item_num}')
                        else:
                            is_list_item = False
                    else:
                        is_list_item = False
                else:
                    is_list_item = False

                # 【WPS智能目录匹配】只提取Heading样式的段落
                # WPS智能目录规则：只显示使用了Heading样式的章节，忽略Normal样式
                if level is not None:
                    is_heading_style = (
                        "heading" in style_name.lower() or "标题" in style_name.lower()
                    )
                    # 如果有大纲级别，说明是TOC中的项，保留
                    has_outline_level = False
                    try:
                        if paragraph._element.pPr is not None:
                            has_outline_level = (
                                paragraph._element.pPr.outlineLvl is not None
                            )
                    except:
                        pass

                    # 过滤条件：非Heading样式 且 没有大纲级别
                    if not is_heading_style and not has_outline_level:
                        # print(
                        #     f"[STYLE-FILTER] 过滤非Heading样式: {text[:50]} (样式: {style_name})"
                        # )
                        continue

                # 增加智能层级跳变约束 (类似于 WPS 的“跃迁保护”)
                # 如果当前已经进入深层章节（Level 2 及以上），突然出现一个 Level 1 标题
                # 且没有明确的标题样式或核心关键字，极大概率是文档内的列表项（如“1. 步骤提示”），应予拦截
                if level == 1 and current_section["level"] >= 2:
                    is_explicit = (
                        "heading" in style_name.lower() or "标题" in style_name.lower()
                    )
                    is_core = any(kw in text for kw in core_keywords)

                    if not (is_explicit or is_core):
                        # 如果不是显式标题样式，也不是核心关键字，则需要极度谨慎
                        has_num_prop = False
                        try:
                            has_num_prop = paragraph._element.pPr.numPr is not None
                        except:
                            pass

                        # 特别策略：即便有自动编号属性，但如果不加粗，在深层嵌套下极大概率是列表而非章节
                        is_bold = any(
                            run.bold
                            for run in paragraph.runs
                            if run.text.strip() and run.bold
                        )
                        if not is_bold:
                            try:
                                if paragraph.style.font.bold:
                                    is_bold = True
                            except:
                                pass

                        # 如果以上都不满足，但也匹配了明显的 "数字. 标题" 模式（且长度很短），
                        # 仍有可能是新章节。
                        is_likely_chapter = False
                        m_tmp = re.match(
                            r"^(\d+)[\.．]\s?([\u4e00-\u9fa5]{2,10})$", text.strip()
                        )
                        if m_tmp:
                            # 提取数字
                            try:
                                new_num = int(m_tmp.group(1))
                                # 核心逻辑：如果识别出的顶级编号（如 2.）小于当前已达到的顶级章节号（如 4.），
                                # 那么它绝对是文档内的列表项，而非新章起始。
                                if new_num <= level_counters[1]:
                                    is_likely_chapter = False
                                else:
                                    # 只有当它是紧接在后的新章（如 1->2 或者当前是 4，新来的是 5）时才信任
                                    is_likely_chapter = True
                            except:
                                pass

                        if has_num_prop:
                            if not (is_bold or is_likely_chapter):
                                # 列表项概率更高
                                level = None
                        else:
                            # 既无 numPr 也无显式样式和核心词，极大概率是普通列表文字
                            if not (is_bold or is_likely_chapter):
                                level = None

                if level:
                    # 【快速路径】如果段落有outline level，应该优先信任outline level
                    # 直接根据outline level生成编号，不必等待TOC匹配
                    has_outline_level_for_numbering = False
                    outline_level_value = None
                    try:
                        pPr = paragraph._element.pPr
                        if pPr is not None and pPr.outlineLvl is not None:
                            outline_level_value = int(pPr.outlineLvl.val) + 1
                            has_outline_level_for_numbering = True
                    except:
                        pass

                    # 检查是否进入"过程说明"章节（对于非TOC匹配的情况）
                    text_without_number = re.sub(r"^[\d\.]+\s*", "", text).strip()
                    procedure_keywords = [
                        "过程说明",
                        "功能过程",
                        "过程描述",
                        "功能过程说明",
                    ]
                    if any(
                        keyword in text_without_number for keyword in procedure_keywords
                    ):
                        if not in_procedure_section:
                            in_procedure_section = True
                            procedure_parent_level = level
                            list_item_counter = 0
                            log_info(f'[PROCEDURE] 进入过程说明章节(非TOC): {text_without_number}, Level={level}')

                    # 检查是否退出过程说明章节（遇到正式章节）
                    elif in_procedure_section and level <= procedure_parent_level:
                        in_procedure_section = False
                        list_item_counter = 0
                        log_info(f'[PROCEDURE] 退出过程说明章节（遇到同级或更高级章节）')

                    # 同步计数器与编号补全
                    clean_title = text
                    # 增强正则表达式：支持 [4.1.1.1] 这种深度嵌套，支持带括号，支持多级
                    manual_match = re.match(
                        r"^\s*[\[\(]?(\d+([\.\．\s]*[\.．]\s*\d+)*)[\]\)]?[\.．\s]*",
                        text,
                    )

                    if manual_match:
                        # 情况 A: 文本自带明确的多级编号 (如 [1.1.1] 或 4.1.1.1)
                        raw_num_str = manual_match.group(1)
                        full_num_str = raw_num_str.replace(" ", "").replace("．", ".")
                        parts = [p for p in full_num_str.split(".") if p.strip()]
                        num_level = len(parts)

                        # 【修复】如果只匹配到单个数字（如"4A门户"中的"4"），
                        # 且原文本后面紧跟非数字字符（如字母），则不视为有效的手动编号
                        # 继续后面的自动编号流程
                        if num_level == 1:
                            # 检查匹配的数字后面是什么字符
                            after_num = text[len(manual_match.group(0)) :]
                            if (
                                after_num
                                and not after_num[0].isdigit()
                                and after_num[0] not in (".", "．", " ", "\t")
                            ):
                                # 单数字后面不是数字或点号，不是有效的章节编号
                                # 例如："4A门户探针拨测"，"4abc"等
                                # 重置manual_match，走自动编号流程
                                manual_match = None

                    if manual_match:
                        # 重新获取编号（因为可能在上面被重置了）
                        raw_num_str = manual_match.group(1)
                        full_num_str = raw_num_str.replace(" ", "").replace("．", ".")
                        parts = [p for p in full_num_str.split(".") if p.strip()]
                        num_level = len(parts)

                        # 核心修复：手动编号严格覆盖初步探测的 level
                        level = num_level

                        # 【关键优化】如果有outline level，优先使用outline level
                        if has_outline_level_for_numbering and outline_level_value:
                            level = outline_level_value
                            log_info(f'[OUTLINE-PRIORITY] 使用outline level覆盖手动编号检测: 文本编号={num_level} → outline_level={outline_level_value}')

                        # 【重复检测】基于(编号,标题)完全匹配检测重复
                        is_duplicate = False
                        # 先提取标题文本
                        matched_str_temp = manual_match.group(0)
                        clean_content_temp = text[len(matched_str_temp) :].strip()
                        if clean_content_temp.startswith(("、", ".", "．", " ")):
                            clean_content_temp = clean_content_temp.lstrip(
                                "、.． "
                            ).strip()

                        chapter_key = (full_num_str, clean_content_temp)
                        if chapter_key in seen_chapters:
                            log_warn(f'[WARN] 检测到重复章节(编号+标题完全相同)，跳过: [{full_num_str}] {clean_content_temp[:40]}')
                            is_duplicate = True

                        if is_duplicate:
                            continue

                        # 记录这个章节
                        seen_chapters.add(chapter_key)

                        if level:
                            # 提取内容：去掉手动编号前缀（包括它后面的空格和点）
                            matched_str = manual_match.group(0)
                            clean_content = text[len(matched_str) :].strip()
                            if clean_content.startswith(("、", ".", "．", " ")):
                                clean_content = clean_content.lstrip("、.． ").strip()

                            # 【关键优化】如果有outline level，用outline level来同步计数器
                            if has_outline_level_for_numbering and outline_level_value:
                                # 基于outline level同步计数器
                                # 只增加该层级，保留上层的值，重置下层
                                level_counters[outline_level_value] += 1
                                for deeper in range(
                                    outline_level_value + 1, len(level_counters)
                                ):
                                    level_counters[deeper] = 0
                                log_info(f'[OUTLINE-SYNC] 基于outline level同步计数器: level={outline_level_value}, counters[1-6]={level_counters[1:7]}')
                            else:
                                # 原有的手动编号同步逻辑
                                # 强制同步计数器
                                for i, p in enumerate(parts):
                                    idx = i + 1
                                    if idx < len(level_counters):
                                        try:
                                            level_counters[idx] = int(p)
                                        except:
                                            pass
                                # 重置子级
                                for i in range(num_level + 1, len(level_counters)):
                                    level_counters[i] = 0

                            # 保留原文档编号格式，不强制添加"."
                            # 提取原始编号后的分隔符
                            sep_match = re.match(
                                rf"^[\s\[\(]*{re.escape(full_num_str)}[\]\)]*([^\u4e00-\u9fa5]*)",
                                text,
                            )
                            if sep_match and len(sep_match.group(1).strip()) > 0:
                                separator = sep_match.group(
                                    1
                                ).rstrip()  # 保留原始分隔符（去掉尾部空格）
                                if (
                                    not separator
                                ):  # 如果分隔符为空，说明编号直接连着中文
                                    separator = ""
                                clean_title = (
                                    f"{full_num_str}{separator}{clean_content}"
                                )
                            else:
                                # 编号直接连着中文，无分隔符
                                clean_title = f"{full_num_str}{clean_content}"

                            # 记录已使用的编号
                            used_numbers.add(full_num_str)

                            log_info(f'[AUTO-NUM] Text: {text[:40]}, ManualNum: {full_num_str}, Level: {level}')
                            log_info(f'    → Using manual numbering, Updated Counters[1-6]: {level_counters[1:7]}')
                        else:
                            # 如果手动编号判定为 None (已经在 get_level_enhanced 中拦截)
                            continue
                    else:
                        # 情况 B: 没有多级手动编号
                        # 【恢复自动编号】根据层级自动生成编号，和WPS显示一致

                        # 特殊处理：列表项
                        if is_list_item:
                            # 【边界检查】防止parent level过大
                            if procedure_parent_level >= len(level_counters):
                                log_warn(f'[WARN] 列表项父层级 {procedure_parent_level} 过深，跳过: {text[:40]}')
                                continue

                            # 列表项：使用父章节编号 + 列表项序号
                            # 例如：父章节是 4.1.1.1.3，列表项1 -> 4.1.1.1.3.1
                            parent_parts = [
                                level_counters[i]
                                for i in range(1, procedure_parent_level + 1)
                            ]
                            parent_num = ".".join(map(str, parent_parts))
                            auto_num = f"{parent_num}.{list_item_counter}"

                            log_info(f'[AUTO-NUM-LIST] ListItem: {text[:40]}, Parent Level: {procedure_parent_level}, Item#: {list_item_counter}')
                            log_info(f'    → Generated: {auto_num}, Counters[1-6]: {level_counters[1:7]}')

                            # 提取列表项文本（去掉"1. "前缀）
                            list_item_match = re.match(r"^\d+[.．]\s+(.+)$", text)
                            if list_item_match:
                                item_text = list_item_match.group(1)
                                clean_title = f"{auto_num} {item_text}"
                            else:
                                clean_title = f"{auto_num} {text}"
                        else:
                            # 【恢复自动编号】普通章节：根据层级生成编号

                            # 【关键修复】如果有outline level且层级>3，说明是深层项目
                            # 这种情况下，应该谨慎处理，不能简单地递增当前层级计数器
                            if (
                                has_outline_level_for_numbering
                                and outline_level_value
                                and outline_level_value > 3
                            ):
                                # 深层项目：尝试从文本中推导编号
                                # 例如"4.1.1.1 xxx"格式
                                text_num_match = re.match(
                                    r"^(\d+(?:\.\d+)+)", text.strip()
                                )
                                if text_num_match:
                                    # 从文本中成功提取了编号
                                    inferred_num = text_num_match.group(1)
                                    parts_inferred = inferred_num.split(".")

                                    # 同步计数器
                                    for i, p in enumerate(parts_inferred):
                                        idx = i + 1
                                        if idx < len(level_counters):
                                            try:
                                                level_counters[idx] = int(p)
                                            except:
                                                pass
                                    # 重置子级
                                    for i in range(
                                        len(parts_inferred) + 1, len(level_counters)
                                    ):
                                        level_counters[i] = 0

                                    log_info(f'[DEEP-OUTLINE] 从文本推导深层编号: {inferred_num}, level={outline_level_value}')
                                else:
                                    # 从文本中找不到编号，但有outline level
                                    # 这种情况较少见，记录警告
                                    log_warn(f"[WARN] 深层项目但未找到编号: '{text[:40]}', outline_level={outline_level_value}")
                                    # 不使用自动编号，而是保留原文
                                    clean_title = text
                                    current_section = {
                                        "level": outline_level_value,
                                        "title": clean_title,
                                        "num": "",
                                        "content": [],
                                        "full_path": "",
                                    }
                                    continue

                            # 检查是否为有效层级（level > 1 时需要检查父级是否存在）
                            if level > 1 and level_counters[level - 1] == 0:
                                # 父级层级为0，这是孤儿章节，可能需要跳过或保留原文
                                # 但为了和WPS一致，这里仍然生成编号
                                pass

                            # 【边界检查】
                            if level >= len(level_counters):
                                log_warn(f'[WARN] 层级 {level} 过深(超过{len(level_counters)}层)，跳过: {text[:40]}')
                                continue

                            # 增加当前层级计数器
                            level_counters[level] += 1

                            # 重置所有更深层级的计数器
                            for deeper in range(level + 1, len(level_counters)):
                                level_counters[deeper] = 0

                            # 构建编号（例如：4.1.2.3）
                            # 【修复】跳过值为0的中间层级，避免生成 3.0.1 这样的编号
                            # 当文档结构跳级时（如 Level 1 → Level 3），只保留有值的层级
                            active_parts = [
                                level_counters[i]
                                for i in range(1, level + 1)
                                if level_counters[i] > 0
                            ]
                            auto_num = ".".join(map(str, active_parts))

                            # 【关键检查】检查生成的编号是否已被使用
                            # 如果已被使用，说明计数器逻辑有问题，需要调整
                            if auto_num in used_numbers:
                                log_warn(f"[WARN] 自动生成的编号 '{auto_num}' 已被使用！可能是计数器逻辑错误")                                # 尝试找到下一个可用的编号
                                counter = 1
                                while (
                                    f"{auto_num.rsplit('.', 1)[0]}.{counter}"
                                    in used_numbers
                                    and counter < 100
                                ):
                                    counter += 1
                                if counter < 100:
                                    auto_num = f"{auto_num.rsplit('.', 1)[0]}.{counter}"
                                    log_info(f'    → 调整为: {auto_num}')

                            log_info(f'[AUTO-NUM] Text: {text[:40]}, Level: {level}')
                            log_info(f'    → Generated: {auto_num}, Counters[1-6]: {level_counters[1:7]}')

                            # 清理标题文本（去除开头的数字编号，避免重复）
                            clean_content_temp = re.sub(
                                r"^[\d\.．]+\s*", "", text
                            ).strip()
                            clean_title = f"{auto_num} {clean_content_temp}"

                            # 记录已使用的编号
                            used_numbers.add(auto_num)

                    # 孤儿目录过滤（仅对自动编号生效）
                    if not manual_match:
                        if level > 1 and level_counters[1] == 0:
                            continue

                    # 【控制台调试输出】打印实际获取到的 Word 目录项及其层级
                    if sections_found < 20:  # 只打印前20个章节的详情
                        log_info(f'[DEBUG-WORD-STRUCTURE] Level: {level} | FinalTitle: {clean_title}')

                    sections_found += 1

                    # 保存前一个章节
                    if (
                        current_section["title"] != "前言/未归类"
                        or current_section["content"]
                    ):
                        current_section["content"] = "\n".join(
                            current_section["content"]
                        ).strip()
                        sections.append(current_section)
                        log_info(f"[APPEND-SECTION #{len(sections)}] Level={current_section['level']}, Title={current_section['title'][:60]}")

                    # 更新路径跟踪
                    current_titles[level] = clean_title
                    for l in range(level + 1, 10):
                        current_titles[l] = ""

                    path_parts = [
                        current_titles[l]
                        for l in range(1, level + 1)
                        if current_titles[l]
                    ]

                    # 提取编号(支持数字后的空格)
                    num_match = re.match(r"^([\d\.]+)", clean_title)
                    section_num = num_match.group(1).rstrip(".") if num_match else ""

                    current_section = {
                        "level": level,
                        "title": clean_title,
                        "num": section_num,
                        "content": [],
                        "full_path": " > ".join(path_parts),
                    }
                else:
                    current_section["content"].append(text)

            # 最后一个章节
            if current_section and (
                current_section["content"] or current_section["level"] > 0
            ):
                current_section["content"] = "\n".join(
                    current_section["content"]
                ).strip()
                sections.append(current_section)

            # 【可选】如果仍有缺失的中间层级，补充占位符
            # 注意：如果文档有toc目录，这一步通常不需要
            # sections = DocumentProcessor._fill_missing_levels(sections)

            # 输出详细的处理统计
            final_processing_end = time.time()
            total_processing_time = final_processing_end - main_processing_start
            overall_time = final_processing_end - start_time
            log_info(f'[DOC-PROCESS] ✓ 第二轮处理完成')
            # print(
            #     f"[DOC-PROCESS] 统计: 总章节={len(sections)}, 添加内容行={content_lines_added}, 第二轮耗时={total_processing_time:.2f}秒, 总耗时={overall_time:.2f}秒"
            # )

            return sections
        except Exception as e:
            import traceback

            log_error(f'提取 Word 结构失败: {e}')
            traceback.print_exc()
            return []
        finally:
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass

    @staticmethod
    def _fill_missing_levels(sections):
        """
        补充缺失的中间层级
        例如：如果有 1 和 4.1.1.1，自动补充 4.1 和 4.1.1
        会尝试从文档的目录(toc)样式段落中查找真实标题
        """
        if not sections:
            return sections

        # 第一步：构建已存在的编号集合
        existing_numbers = set()
        for section in sections:
            match = re.match(r"^(\d+(?:\.\d+)*)\.\s+", section["title"])
            if match:
                existing_numbers.add(match.group(1))

        # 第二步：收集需要补充的中间层级
        missing_levels = {}  # {编号: depth}

        for section in sections:
            title = section["title"]
            match = re.match(r"^(\d+(?:\.\d+)*)\.\s+(.+)$", title)
            if not match:
                continue

            full_number = match.group(1)
            parts = full_number.split(".")

            # 检查所有父级层级
            for i in range(1, len(parts)):
                parent_number = ".".join(parts[:i])
                if (
                    parent_number not in existing_numbers
                    and parent_number not in missing_levels
                ):
                    missing_levels[parent_number] = i

        # 如果没有缺失层级，直接返回
        if not missing_levels:
            return sections

        # 第三步：为缺失层级创建占位章节
        filled_sections = list(sections)
        for number, depth in missing_levels.items():
            filled_sections.append(
                {
                    "level": depth,
                    "title": f"{number}. (层级占位)",
                    "content": "",
                    "full_path": "",
                }
            )
            log_info(f'[FILL-GAP] 补充缺失层级: {number}. (层级占位)')

        # 第四步：按编号排序
        def sort_key(s):
            match = re.match(r"^(\d+(?:\.\d+)*)", s["title"])
            if match:
                parts = match.group(1).split(".")
                return [int(p) for p in parts]
            return [0]

        filled_sections.sort(key=sort_key)
        return filled_sections

    @staticmethod
    def get_functional_points_count(file_path, sheet_name=None):
        """
        获取 Excel 中有效功能点总数 (即数据行数)
        - 仅计算子功能过程
        - 跳过空行 (现由于用户反馈建议：改为遇到空行停止)
        - 遇到“注：”等提示文字停止遍历
        """
        try:
            import openpyxl
            from utils.runtime_logger import RuntimeLogger

            # 检查是否已是 Workbook 对象
            if hasattr(file_path, "sheetnames") and hasattr(file_path, "active"):
                # 这是一个 Workbook 对象
                wb = file_path
            else:
                # 增加 read_only=True 提高大型/复杂 XML 文件的加载成功率
                # 这种模式比普通加载占用内存更少，且对某些特殊格式（如超长行）更鲁棒
                try:
                    wb = openpyxl.load_workbook(
                        file_path, data_only=True, read_only=True
                    )
                except Exception as e:
                    RuntimeLogger.log(
                        f"Excel 初始加载失败: {e}，尝试标准模式加载...", level="WARN"
                    )
                    wb = openpyxl.load_workbook(file_path, data_only=True)

            target_sheet_name = sheet_name

            if not target_sheet_name:
                for name in wb.sheetnames:
                    if "功能点拆分" in name:
                        target_sheet_name = name
                        break

            if not target_sheet_name or target_sheet_name not in wb.sheetnames:
                # 兜底：如果找不到且未指定，尝试第一个工作表
                if not sheet_name and wb.sheetnames:
                    target_sheet_name = wb.sheetnames[0]
                else:
                    return 0

            ws = wb[target_sheet_name]

            # 1. 探测表头区域以确定数据起点
            header_end = -1
            target_col = -1
            # read_only 模式下使用 iter_rows 效率最高
            probe_rows = list(
                ws.iter_rows(min_row=1, max_row=40, min_col=1, max_col=20)
            )
            for row_idx, row_cells in enumerate(probe_rows, 1):
                row_vals = [str(cell.value) for cell in row_cells]
                row_str = " ".join([v for v in row_vals if v != "None"])

                # 常见的关键字段标记表头结束
                if any(
                    kw in row_str
                    for kw in ["客户需求", "一级模块", "功能过程", "子过程", "功能描述"]
                ):
                    header_end = row_idx
                    # 确定“子过程”所在的列
                    for col_idx, cell in enumerate(row_cells, 1):
                        val = str(cell.value or "")
                        if "子过程" in val:
                            target_col = col_idx
                            break
                    if target_col == -1:
                        for col_idx, cell in enumerate(row_cells, 1):
                            val = str(cell.value or "")
                            if "功能过程" in val:
                                target_col = col_idx
                                break
                    break

            if header_end == -1:
                header_end = 0  # 没有探测到则从头开始

            # 2. 遍历计数
            count = 0
            # 使用 iter_rows 从 header_end+1 开始遍历
            for r_idx, row in enumerate(
                ws.iter_rows(min_row=header_end + 1), header_end + 1
            ):
                # 检查是否触底 (图 3 中的注记文字)
                row_note_found = False
                instructional_row = False

                # 预提取前几个单元格的字符串，用于检测说明文字
                row_head_texts = []
                for i in range(min(5, len(row))):
                    val = str(row[i].value or "").strip()
                    row_head_texts.append(val)

                    if (
                        val.startswith("注：")
                        or val.startswith("注:")
                        or "请在正式提交时删除" in val
                    ):
                        row_note_found = True
                        break

                    # 检查是否是模板说明文字 (Instructional Text)
                    # 这里的逻辑与 HierarchicalMatcher._is_instructional_text 保持一致
                    if val:
                        clean_v = re.sub(r"\s+", "", val)
                        for pattern in HierarchicalMatcher.IGNORE_PATTERNS:
                            if re.search(pattern.replace(" ", ""), clean_v):
                                instructional_row = True
                                break
                    if instructional_row:
                        break

                if row_note_found:
                    break

                # 如果是说明提示行，跳过本行不计数
                if instructional_row:
                    continue

                # 检查整行是否为空 (探测前15列)
                row_has_data = False
                for i in range(min(15, len(row))):
                    if row[i].value is not None:
                        row_has_data = True
                        break

                if not row_has_data:
                    # 根据用户建议：整行为空就可以停止计数
                    break

                # 如果找到了目标列，则该列不为空才计数
                if target_col != -1 and target_col <= len(row):
                    if row[target_col - 1].value is not None:
                        count += 1
                else:
                    count += 1

            if hasattr(wb, "close"):
                wb.close()
            return count
        except Exception as e:
            try:
                from utils.runtime_logger import RuntimeLogger

                RuntimeLogger.log(f"获取功能点数失败: {e}", level="ERROR")
            except:
                pass
            log_error(f'获取功能点数失败: {e}')
            return 0

    @staticmethod
    def check_adjustment_factors_in_word(file_path, target_sections=None):
        """
        深度扫描 Word 正文段落及表格中的附加值调整因子 (Node 4)
        支持文件路径（str）或已加载的 Document 对象
        """
        temp_docx = None
        doc = None
        try:
            from docx import Document
            from docx.text.paragraph import Paragraph
            from docx.table import Table
            import re

            # 检查是否已是 Document 对象
            if hasattr(file_path, "paragraphs") and hasattr(file_path, "element"):
                # 这是一个 Document 对象
                doc = file_path
            else:
                # 兼容 pathlib.Path
                file_path_str = str(file_path)
                # .doc 转换处理
                if file_path_str.lower().endswith(".doc"):
                    temp_docx = DocumentProcessor._convert_doc_to_docx(file_path_str)
                    if not temp_docx:
                        return {}
                    doc_to_read = temp_docx
                else:
                    doc_to_read = file_path_str

                doc = Document(doc_to_read)

            # 因子初始化
            factors = {
                "scale": {
                    "name": "需求变更规模因子",
                    "value": None,
                    "text_value": None,
                    "table_value": None,
                    "found_in_text": False,
                    "found_in_table": False,
                },
                "distributed": {
                    "name": "分布式处理",
                    "value": None,
                    "text_value": None,
                    "table_value": None,
                    "found_in_text": False,
                    "found_in_table": False,
                },
                "performance": {
                    "name": "性能",
                    "value": None,
                    "text_value": None,
                    "table_value": None,
                    "found_in_text": False,
                    "found_in_table": False,
                },
                "reliability": {
                    "name": "可靠性",
                    "value": None,
                    "text_value": None,
                    "table_value": None,
                    "found_in_text": False,
                    "found_in_table": False,
                },
                "multiple_sites": {
                    "name": "多重站点",
                    "value": None,
                    "text_value": None,
                    "table_value": None,
                    "found_in_text": False,
                    "found_in_table": False,
                },
            }

            # --- [NEW] 范围限定逻辑：基于目录树精确定位起止路标 ---
            adj_parent_found = False
            scale_info = None
            quality_info = None

            RuntimeLogger.log(
                f"   [Node 4] 开始附加值因子提取 (Word)，提供目录节点数: {len(target_sections) if target_sections else 0}"
            )

            if target_sections is None:
                # 兜底：如果没传，内部提取一次
                target_sections = DocumentProcessor.extract_word_structure(doc)

            for s in target_sections:
                t = (s.get("title") or "").strip()
                # [极简匹配] 只要包含关键字就锁定章节
                t_lower = t.lower()
                if any(x in t_lower for x in ["附加值", "调整因子", "value factors"]):
                    adj_parent_found = True
                    RuntimeLogger.log(f"   [Node 4] 找到章节父级: {t}")
                if (
                    any(x in t_lower for x in ["规模因子", "scale factor"])
                    and not scale_info
                ):
                    scale_info = s
                if (
                    any(
                        x in t_lower
                        for x in [
                            "质量及特性",
                            "特性需求",
                            "quality requirement",
                            "factor description",
                        ]
                    )
                    and not quality_info
                ):
                    quality_info = s

            # 收集特定范围内的块
            scale_paras, scale_tables = [], []
            quality_paras, quality_tables = [], []
            quality_keys = [
                "distributed",
                "performance",
                "reliability",
                "multiple_sites",
            ]

            if adj_parent_found:
                active_category = None
                scale_lvl = scale_info.get("level") if scale_info else 0
                quality_lvl = quality_info.get("level") if quality_info else 0

                RuntimeLogger.log(
                    f"   [Node 4] 章节定位: 规模因子段={scale_info.get('title') if scale_info else '未找到'}, 质量因子段={quality_info.get('title') if quality_info else '未找到'}"
                )

                # [性能修复] 目录标题只规范化一次，建哈希表做精确匹配。
                # 原实现对"每个短段落 × 每个目录节点"都重复执行 re.sub，
                # O(N×M) 在大文档上是数百万次正则调用，仅此一步就耗时 3 分钟以上。
                section_lookup = {}
                for s in target_sections:
                    s_title = (s.get("title") or "").strip()
                    tmp_s = s_title.lstrip(". \t\n\r")
                    core_s = re.sub(
                        r"^[^\w\u4e00-\u9fa5]*\d+[\.\d\s、-]*", "", tmp_s
                    ).strip()
                    if core_s and core_s not in section_lookup:
                        section_lookup[core_s] = s

                # 利用 doc.element.body.iterchildren() 确保正文和表格在收集时保持原始顺序
                for child in doc.element.body.iterchildren():
                    is_p = child.tag.endswith("p")
                    is_tbl = child.tag.endswith("tbl")

                    if is_p:
                        p_obj = Paragraph(child, doc)
                        txt = p_obj.text.strip()
                        # 降低标题识别长度上限 (从 200 降至 100)，避免长句误触发
                        if txt and len(txt) < 100:
                            # [核心优化]：去除编号后与目录树进行比对锁定位置
                            tmp_txt = txt.lstrip(". \t\n\r")
                            clean_txt = re.sub(
                                r"^[^\w\u4e00-\u9fa5]*\d+[\.\d\s、-]*", "", tmp_txt
                            ).strip()
                            matched_s = None

                            if clean_txt and len(clean_txt) < 50:
                                matched_s = section_lookup.get(clean_txt)

                            if matched_s:
                                last_cat = active_category
                                # [FIX] 只有匹配到非目标节点时，才执行退出逻辑
                                if matched_s == scale_info:
                                    active_category = "scale"
                                elif matched_s == quality_info:
                                    active_category = "quality"
                                else:
                                    # 层级退出逻辑：进入了与目标章节同级或更高级的其他章节
                                    # [优化] 增加父子级关系判定：如果是子章节，不自动退出
                                    is_child = False
                                    if active_category == "scale" and scale_info:
                                        # 如果新章节的编号是以父章节编号开头的，视为子章节
                                        if matched_s.get(
                                            "parent_num"
                                        ) == scale_info.get("num") or matched_s.get(
                                            "num", ""
                                        ).startswith(
                                            scale_info.get("num", "NEVERMATCH") + "."
                                        ):
                                            is_child = True
                                    elif active_category == "quality" and quality_info:
                                        if matched_s.get(
                                            "parent_num"
                                        ) == quality_info.get("num") or matched_s.get(
                                            "num", ""
                                        ).startswith(
                                            quality_info.get("num", "NEVERMATCH") + "."
                                        ):
                                            is_child = True

                                    if not is_child:
                                        if (
                                            active_category == "scale"
                                            and matched_s["level"] <= scale_lvl
                                        ):
                                            active_category = None
                                        elif (
                                            active_category == "quality"
                                            and matched_s["level"] <= quality_lvl
                                        ):
                                            active_category = None

                                if last_cat != active_category:
                                    RuntimeLogger.log(
                                        f"   [Node 4] 扫描区域切换: {last_cat} -> {active_category} (由此段触发: '{txt[:30]}...')"
                                    )

                    # 收集数据块：此处不受 200 字符限制，只要在 active_category 范围内即收集
                    if active_category == "scale":
                        if is_p:
                            scale_paras.append(p_obj)
                        elif is_tbl:
                            scale_tables.append(Table(child, doc))
                    elif active_category == "quality":
                        if is_p:
                            quality_paras.append(p_obj)
                        elif is_tbl:
                            quality_tables.append(Table(child, doc))

            RuntimeLogger.log(
                f"   [Node 4] 收集区域统计: 规模因子({len(scale_paras)}段/{len(scale_tables)}表), 质量因子({len(quality_paras)}段/{len(quality_tables)}表)"
            )

            # 1. 扫描段落 (正文文本查找)
            # 兼容带引号(中文/英文)和不带引号的情况
            # 扩充句式：支持 "目前属于", "按照", "处于", "属于" 等更多描述方式
            scale_regex = re.compile(
                r'(?:目前属于|按照|处于|属于|编制为)\s*[“"「「]?\s*(结算|预算|概算|匡算)\s*[”"」」]?\s*阶段'
            )

            # A. 扫描规模因子段落 (限定范围)
            for para in scale_paras:
                text = para.text.strip()
                if not text:
                    continue

                # A. 检查规模因子 (特定句式优先)
                m_scale = scale_regex.search(text)
                if m_scale:
                    val = m_scale.group(1).strip()
                    factors["scale"]["found_in_text"] = True
                    factors["scale"]["text_value"] = val
                    factors["scale"]["value"] = val
                    RuntimeLogger.log(
                        f"   [Node 4] 提取到规模因子(文): {val} (来自: '{text}')"
                    )

                # 备用：传统模糊匹配 (只要段落包含“规模变更因子”关键词即可)
                elif ("规模变更因子" in text or "规模因子" in text) and not factors[
                    "scale"
                ]["text_value"]:
                    m = re.search(r"(结算|预算|概算|匡算)", text)
                    if m:
                        val = m.group(1)
                        factors["scale"]["found_in_text"] = True
                        factors["scale"]["text_value"] = val
                        factors["scale"]["value"] = val
                        RuntimeLogger.log(f"   [Node 4] 提取到规模因子(文-模糊): {val}")

            # B. 扫描质量特性段落 (限定范围)
            last_key = None
            for para in quality_paras:
                text = para.text.strip()
                if not text:
                    continue

                # 特殊情况处理：如果整段话就是“无”，且没有具体因子关键字，可能意味着所有因子均为默认
                if text == "无" or text == "无要求":
                    factors["_global_default"] = True
                    RuntimeLogger.log(f"   [Node 4] 发现全局默认(无)标记")

                found_any_key_in_this_para = False
                for key in [
                    "distributed",
                    "performance",
                    "reliability",
                    "multiple_sites",
                ]:
                    kw = factors[key]["name"]
                    # 优化关键字匹配：支持中文数字编号 (一、二、三、四)
                    is_key_match = False
                    if text.startswith(kw) or re.search(
                        rf"^[①-⑩\d\s\.、\(\)（）一二三四五六七八九十]*{kw}", text
                    ):
                        # 确保是标题行（后面带冒号或整体较短）
                        if re.search(rf"{kw}[:：\s]", text) or len(text) < 40:
                            is_key_match = True

                    if is_key_match:
                        # 排除目录行
                        if text.count(".") > 5:
                            continue

                        last_key = key
                        found_any_key_in_this_para = True
                        factors[key]["found_in_text"] = True
                        RuntimeLogger.log(
                            f"   [Node 4] 匹配到因子关键字: {kw} (原文: '{text[:40]}...')"
                        )

                        # 1. 查找明确的数字 (仅限 -1, 0, 1，且排除 7*24 干扰)
                        # 查找格式：关键字 + 冒号/空格 + (-1/0/1)
                        m_num = re.search(rf"{kw}.*?[:：]\s*(-?1|0)\b(?![\*x])", text)
                        if m_num:
                            val = m_num.group(1)
                            # [USER UPDATE] 0 视为 1 (有描述)
                            final_val = "1" if val == "0" else val
                            factors[key]["text_value"] = final_val
                            factors[key]["value"] = final_val
                            RuntimeLogger.log(
                                f"   [Node 4] 提取到明确分值 {kw}: {val} -> 归一化为 {final_val}"
                            )
                            continue

                        # 2. 查找明确的“空”字
                        if re.search(rf"{kw}.*?[:：]?\s*空", text):
                            factors[key]["text_value"] = "空"
                            factors[key]["value"] = "空"
                            continue

                        # 3-5. 查找预定义短语
                        negative_phrases = {
                            "distributed": [
                                "没有明示对分布式处理的需求事项",
                                "无分布式",
                                "无说明",
                                "-1",
                            ],
                            "performance": [
                                "没有明示对性能的特别需求事项或仅需提供基本性能",
                                "仅需提供基本性能",
                                "无明示对性能的特别需求",
                                "-1",
                            ],
                            "reliability": [
                                "没有明示对可靠性的特别需求事项或仅需提供基本的可靠性",
                                "仅需提供基本的可靠性",
                                "无明示对可靠性的特别需求",
                                "-1",
                            ],
                            "multiple_sites": [
                                "在相同用途的硬件或软件环境下运行",
                                "同一套硬件环境",
                                "无站点差异",
                                "-1",
                            ],
                        }

                        # 补充通用的否定词 (作为备选)
                        gen_neg = ["无", "不涉及", "无要求", "不适用", "为空", "此项无"]

                        found_val = None
                        # [新逻辑] 首先假设找到了(因为进到了关键字段落)
                        factors[key]["found_in_text"] = True

                        # 检查具体的负面完整长短语 (高优先级)
                        is_neg_matched = False
                        for phrase in negative_phrases.get(key, []):
                            if phrase in text:
                                is_neg_matched = True
                                break

                        # 检查通用的单字/短词否定
                        if not is_neg_matched:
                            # 匹配格式：关键字 + 冒号 + 否定词 (例如 "性能: 无")
                            for gn in gen_neg:
                                if re.search(rf"{kw}[:：\s]*{gn}\b", text) or (
                                    len(text) < 15 and gn in text
                                ):
                                    is_neg_matched = True
                                    break

                        if is_neg_matched:
                            factors[key]["found_in_text"] = False
                            factors[key]["text_value"] = "缺失"  # 标记为业务上的缺失
                            factors[key]["value"] = "-1"
                            found_val = "-1"
                            RuntimeLogger.log(
                                f"   [Node 4] 匹配到负向/缺失描述: {kw} -> 视为缺失(-1)"
                            )
                            continue

                        # [USER UPDATE] 文字有写就是1 (只要不是否定描述，且进了这个段落，就视为1)
                        # 排除掉仅含标题名称的极短行
                        clean_text = text.strip()
                        if len(clean_text) > len(kw) + 1:
                            factors[key]["found_in_text"] = True
                            factors[key]["text_value"] = "1"
                            factors[key]["value"] = "1"
                            RuntimeLogger.log(
                                f"   [Node 4] 文字存在有效描述 -> 判定为 1"
                            )
                            continue
                        else:
                            # 如果确实很短，可能是个标题占位，记录但暂不设为1，等后续段落内容
                            factors[key]["found_in_text"] = True
                            factors[key]["text_value"] = "已识别标题"
                            # 暂不设 factors[key]["value"]，让后面的段落内容来填充

                # [Optimization] 粘性上下文：如果当前段落没有发现新关键字，但存在 last_key，
                # 且当前段落有实质内容，则补全 last_key 的有效性判定
                if not found_any_key_in_this_para and last_key:
                    if len(text) > 5 and not factors[last_key].get("value"):
                        # 只要不是明显的新章节标题，就视为上一个因子的描述延伸
                        if not re.match(r"^[一二三四五六七八九十]、", text):
                            factors[last_key]["text_value"] = "1"
                            factors[last_key]["value"] = "1"
                            factors[last_key]["found_in_text"] = True

                # 针对“均一致”或“均设置为-1”等合并描述
                if ("均" in text or "都" in text) and any(
                    kw in text for kw in ["分布式", "性能", "可靠", "站点"]
                ):
                    target_val = None
                    # [USER UPDATE] 遵循您的最新逻辑：只有明确提到 -1 或 “无” 才当做 -1
                    # 提到 0、1 或 “基本” 均视为有描述 (1)
                    if any(x in text for x in ["-1", "无"]):
                        target_val = "-1"
                    elif any(x in text for x in ["0", "1", "基本", "一致", "有"]):
                        target_val = "1"

                    if target_val:
                        for key in [
                            "distributed",
                            "performance",
                            "reliability",
                            "multiple_sites",
                        ]:
                            # 只有在还没被具体识别的情况下，才用合并描述兜底
                            if (
                                not factors[key].get("value")
                                or factors[key]["value"] == "-1"
                            ):
                                factors[key]["found_in_text"] = target_val == "1"
                                factors[key]["text_value"] = target_val
                                factors[key]["value"] = target_val
                        RuntimeLogger.log(
                            f"   [Node 4] 识别合并描述: '{text}' -> 统一设置为 {target_val}"
                        )

            # 2. 扫描表格
            target_tables = scale_tables + quality_tables
            for table in target_tables:
                all_table_text = "".join([c.text for r in table.rows for c in r.cells])

                # [NEW] 过滤掉模板参考表 (即图1、图2)
                # 仅当表格内容极其像模板定义时才跳过
                if any(
                    x in all_table_text
                    for x in ["判断标准", "没有明示", "分值标识", "附加值因素"]
                ) or (
                    "2.00" in all_table_text
                    and "1.50" in all_table_text
                    and "1.26" in all_table_text
                ):
                    RuntimeLogger.log(
                        "   [Node 4] 跳过表格: 判定为参考模版参考表(图1/图2)。"
                    )
                    continue

                # 规模因子表检测
                is_scale_table = False
                if (
                    "需求变更规模因子" in all_table_text
                    or "变更规模因子" in all_table_text
                ):
                    is_scale_table = True

                # 特性表检测
                found_count = sum(
                    1 for qk in quality_keys if factors[qk]["name"] in all_table_text
                )
                if not is_scale_table and found_count < 1:
                    continue

                RuntimeLogger.log(
                    f"   [Node 4] 扫描表格: 行数={len(table.rows)}, 命中因子数={found_count}"
                )

                for row in table.rows:
                    cells_text = [cell.text.strip() for cell in row.cells]
                    row_content = " ".join(cells_text)

                    if is_scale_table and not factors["scale"]["table_value"]:
                        m = re.search(r"(结算|预算|概算|匡算)", row_content)
                        if m:
                            val = m.group(1)
                            factors["scale"]["found_in_table"] = True
                            factors["scale"]["table_value"] = val
                            if not factors["scale"]["value"]:
                                factors["scale"]["value"] = val
                            RuntimeLogger.log(f"   [Node 4] 表格提取规模因子: {val}")

                    for key in quality_keys:
                        kw = factors[key]["name"]
                        if any(kw in str(c) for c in cells_text[:2]):
                            factors[key]["found_in_table"] = True

                            possible_val = None
                            for cell_text in reversed(cells_text):
                                clean_cell = (
                                    cell_text.strip()
                                    .replace(" ", "")
                                    .replace("\n", "")
                                    .replace("\r", "")
                                )
                                if re.match(r"^-?\d+(\.\d+)?$", clean_cell):
                                    possible_val = clean_cell
                                    break

                            if possible_val:
                                # [USER UPDATE] 表格中 1 和 0 都视为 1，只有 -1 维持 -1
                                if possible_val in ["0", "1"]:
                                    possible_val = "1"

                                factors[key]["table_value"] = possible_val
                                factors[key]["value"] = possible_val
                                RuntimeLogger.log(
                                    f"   [Node 4] 表格提取 {kw}: {possible_val}"
                                )
                            else:
                                # 尝试通过描述（判断标准列）自动计算
                                # [USER UPDATE] 遵循：基本需求描述在该规则下视为 1（因为表格 0 当成 1）
                                # 但“没有明示”这种描述依然维持 -1
                                standards = {
                                    "distributed": {
                                        "-1": "没有明示对分布式处理的需求事项",
                                        "1_basic": "通过网络进行客户端/服务器及网络基础应用分布式处理和传输",
                                        "1": "通过特别的设计保证在多个服务器及处理器上同时相互执行应用中的处理功能",
                                    },
                                    "performance": {
                                        "-1": "没有明示对性能的特别需求事项或仅需提供基本性能",
                                        "1_basic": "存在对连动系统结束处理时间的限制",
                                        "1": "要求设计阶段开始进行性能分析",
                                    },
                                    "reliability": {
                                        "-1": "没有明示对可靠性的特别需求事项或仅需提供基本的可靠性",
                                        "1_basic": "发生故障时带来较多不便或经济损失",
                                        "1": "发生故障时造成重大经济损失或有生命危害",
                                    },
                                    "multiple_sites": {
                                        "-1": "在相同用途的硬件或软件环境下运行",
                                        "1_basic": "在用途类似的硬件或软件环境下运行",
                                        "1": "在不同用途的硬件或软件环境下运行",
                                    },
                                }

                                normalized_row = (
                                    row_content.replace(" ", "")
                                    .replace("\n", "")
                                    .replace("\r", "")
                                )
                                for val, desc in standards.get(key, {}).items():
                                    if desc.replace(" ", "") in normalized_row:
                                        # 将内部标记的 1_basic 统一为 1
                                        final_val = "1" if "1_basic" in val else val
                                        factors[key]["table_value"] = final_val
                                        factors[key]["value"] = final_val
                                        RuntimeLogger.log(
                                            f"   [Node 4] 表格描述提取 {kw}: {final_val}"
                                        )
                                        break
                                        RuntimeLogger.log(
                                            f"   [Node 4] 表格描述提取 {kw}: {val}"
                                        )
                                        break

                                # 如果仍然没提取到具体的数值，但找到了关键字，标记一下
                                if not factors[key].get("table_value"):
                                    factors[key]["table_value"] = "已勾选"

            # 3. 结果汇总与一致性检查
            for key in quality_keys:
                f = factors[key]
                f["is_found"] = f["found_in_text"] or f["found_in_table"]
                f["consistency_warn"] = False

                # [核心判定] 文字与表格取值不一致
                v_text = str(f.get("text_value") or "").strip()
                v_table = str(f.get("table_value") or "").strip()

                # 统一取值归一化：将 "1", "0", "-1" 以外的文本也参与比对
                # 如果一方有具体分值，另一方是“缺失”，则视为不一致
                if f["found_in_text"] != f["found_in_table"]:
                    f["consistency_warn"] = True
                    f["consistency_msg"] = "文字与表格存在缺失项不匹配"
                elif f["found_in_text"] and f["found_in_table"]:
                    if v_text != v_table:
                        f["consistency_warn"] = True
                        f["consistency_msg"] = (
                            f"文字({v_text})与表格({v_table})分值不符"
                        )

            # 补齐默认值
            for key in quality_keys:
                if not factors[key].get("value") or factors[key]["value"] == "缺失":
                    # 只有当彻底没有任何发现时，才设为 -1 缺失
                    if not factors[key]["is_found"]:
                        factors[key]["value"] = "-1"
                        factors[key]["text_value"] = "缺失"
                        factors[key]["table_value"] = "缺失"
                    else:
                        # [USER UPDATE] 只要有一个地方找到了非否定内容，且不是明确的 -1，就视为 1 (显著需求)
                        factors[key]["value"] = "1"

            return factors
        except Exception as e:
            log_error(f'提取调整因子失败: {e}')
            return {}
        finally:
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass

    @staticmethod
    def check_excel_empty_cells(file_path, sheet_name=None, report_base_name=None, project_name=None):

        """
        检查 Excel 的空值情况 (精准合并单元格判定版 V3)

        Args:
            file_path: Excel 文件路径
            sheet_name: 工作表名称（可选，默认自动寻找包含"功能点拆分"的表）
            report_base_name: 统一的报告基础名称（用于提取时间戳）

        Returns:
            dict: {
                'is_ok': bool,
                'errors': list[str],
                'report_path': str
            }
        """
        try:
            import openpyxl
            import pandas as pd
            from utils.initial_review_report import report_file_path

            wb = openpyxl.load_workbook(file_path, data_only=True)

            # 1. 自动定位工作表
            target_sheet_name = sheet_name
            if not target_sheet_name:
                for name in wb.sheetnames:
                    if "功能点拆分" in name:
                        target_sheet_name = name
                        break
            if not target_sheet_name or target_sheet_name not in wb.sheetnames:
                if not sheet_name and wb.sheetnames:
                    target_sheet_name = wb.sheetnames[0]
                else:
                    return {
                        "is_ok": False,
                        "errors": ["未找到有效的工作表进行空值校验"],
                        "report_path": None
                    }

            ws = wb[target_sheet_name]

            # 2. 探测表头区域 (多行探测)
            header_start = -1
            header_end = -1
            for row_idx in range(1, 31):
                row_vals = [str(ws.cell(row=row_idx, column=col).value) for col in range(1, 20)]
                row_str = " ".join([v for v in row_vals if v != "None"])
                if "客户需求" in row_str or "一级模块" in row_str:
                    if header_start == -1: header_start = row_idx
                    header_end = row_idx
            if header_start == -1:
                return {"is_ok": False, "errors": ["未能在工作表中定位到“客户需求”或“一级模块”表头行"],
                        "report_path": None}

            # 3. 定位关键校验列
            target_keywords = [
                "客户需求", "一级模块", "二级模块", "三级模块", "功能用户",
                "触发事件", "功能过程", "子过程描述", "数据移动类型",
                "数据组", "数据属性", "复用度", "CFP"
            ]
            col_map = {}
            for col_idx in range(1, ws.max_column + 1):
                cell_val = ""
                for h_idx in range(header_start, header_end + 1):
                    val = ws.cell(row=h_idx, column=col_idx).value
                    if val: cell_val += str(val)
                for kw in target_keywords:
                    if kw in cell_val:
                        col_map[kw] = col_idx
                        break
            if not col_map:
                return {"is_ok": False, "errors": ["未匹配到任何待校验的关键列，请检查表头名称"], "report_path": None}

            # 4. 确定有效数据范围
            last_valid_row = header_end
            for r in range(header_end + 1, ws.max_row + 1):
                first_cell_val = str(ws.cell(row=r, column=1).value or "").strip()
                if first_cell_val.startswith("注：") or "请在正式提交时删除" in first_cell_val: break
                row_has_something = False
                for c in range(1, 21):
                    if ws.cell(row=r, column=c).value is not None: row_has_something = True; break
                if row_has_something: last_valid_row = r
            if last_valid_row <= header_end: return {"is_ok": True, "errors": [], "report_path": None}

            # 5. 建立合并单元格查询表
            merged_lookup = {}
            for merged_range in ws.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged_range.bounds
                tl_val = ws.cell(row=min_row, column=min_col).value
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1): merged_lookup[(r, c)] = tl_val

            # 6. 遍历扫描空值
            col_errors = {kw: [] for kw in col_map.keys()}
            for r_idx in range(header_end + 1, last_valid_row + 1):
                for kw, c_idx in col_map.items():
                    val = ws.cell(row=r_idx, column=c_idx).value
                    actual_val = val
                    if val is None or str(val).strip() == "" or str(val).lower() == "nan":
                        if (r_idx, c_idx) in merged_lookup: actual_val = merged_lookup[(r_idx, c_idx)]
                    if actual_val is None or str(actual_val).strip() == "" or str(actual_val).lower() == "nan":
                        col_errors[kw].append(r_idx)

            # 7. 格式化错误信息
            final_errors = []
            for kw in target_keywords:
                if kw not in col_errors: continue
                rows = sorted(list(set(col_errors[kw])))
                if not rows: continue
                ranges = []
                start = rows[0];
                prev = start
                for curr in rows[1:]:
                    if curr == prev + 1:
                        prev = curr
                    else:
                        ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                        start = curr;
                        prev = curr
                ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                final_errors.append(f"拆分表中第{', '.join(ranges)}行{kw}为空")

            # 8. 【核心修复】使用初评报告统一路径模块，确保进入本次运行的时间戳文件夹
            report_path = None
            try:
                # 【修改】：优先使用透传下来的 project_name (即 display_name)，如果没有则回退到 file_path
                _, report_path = report_file_path(
                    report_type="Excel空值检查报告",
                    project_name_or_path=file_path,
                    report_base_name=report_base_name,
                )

                report_data = []
                if final_errors:
                    for error_msg in final_errors: report_data.append({"检查项": error_msg})
                else:
                    report_data.append({"检查项": "✅ 所有关键列空值检查通过"})

                df_report = pd.DataFrame(report_data)
                with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
                    df_report.to_excel(writer, index=False, sheet_name="空值检查结果")
            except Exception as e:
                try:
                    from utils.runtime_logger import RuntimeLogger
                    RuntimeLogger.log(f"生成Excel空值检查报告失败: {e}", level="WARN")
                except:
                    pass

            return {"is_ok": len(final_errors) == 0, "errors": final_errors, "report_path": report_path}

        except Exception as e:
            return {"is_ok": False, "errors": [f"Excel 校验引擎异常 (V3): {str(e)}"], "report_path": None}

    @staticmethod
    def extract_excel_info(file_path):
        """
        提取 Excel 的工作表和列信息
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            info = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                info[sheet_name] = df.columns.tolist()
            return info
        except Exception as e:
            log_error(f'提取 Excel 信息失败: {e}')
            return {}

    @staticmethod
    def get_excel_modules(file_path, sheet_name=None):
        """
        提取 Excel 中的三级模块名称集合
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            target_sheet = sheet_name
            if not target_sheet:
                for sheet in excel_file.sheet_names:
                    if "功能点拆分" in sheet:
                        target_sheet = sheet
                        break

            if not target_sheet or target_sheet not in excel_file.sheet_names:
                if not sheet_name and excel_file.sheet_names:
                    target_sheet = excel_file.sheet_names[0]
                else:
                    return []

            # 1. 探测表头区域 (复用探测逻辑)
            df_preview = pd.read_excel(
                file_path, sheet_name=target_sheet, nrows=30, header=None
            )
            header_start = -1
            header_end = -1
            for i, row in df_preview.iterrows():
                row_str = " ".join([str(x) for x in row if pd.notna(x)])
                if "三级模块" in row_str or "一级模块" in row_str:
                    if header_start == -1:
                        header_start = i
                    header_end = i
            if header_start == -1:
                return []

            # 2. 读取数据
            df = pd.read_excel(
                file_path, sheet_name=target_sheet, skiprows=header_start
            )

            # 使用最后一行作为列名
            actual_header_row = df_preview.iloc[header_end]
            final_col_names = [
                str(x) if pd.notna(x) else f"Unnamed_{i}"
                for i, x in enumerate(actual_header_row)
            ]
            df.columns = final_col_names

            # 跳过表头
            df = df.iloc[header_end - header_start + 1 :].reset_index(drop=True)

            if "三级模块" in df.columns:
                return df["三级模块"].ffill().dropna().unique().tolist()
            return []
        except:
            return []

    @staticmethod
    def validate_hierarchy_matching(
            word_path, excel_path, header_row=0, level1_col=None, level2_col=None,
            level3_col=None, sheet_name=None, fuzzy_match=True, threshold=0.8,
            progress_callback=None, word_sections_preloaded=None, project_name=None,
            report_base_name=None,
    ):
        try:
            import pandas as pd
            import os
            from utils.initial_review_report import run_output_dir

            matcher = HierarchicalMatcher(fuzzy_match=fuzzy_match, threshold=threshold)
            if word_sections_preloaded is None:
                if isinstance(word_path, str) and word_path:
                    word_sections_preloaded = DocumentProcessor.extract_word_structure(word_path, use_stable=True)
                elif hasattr(word_path, "paragraphs"):
                    word_sections_preloaded = DocumentProcessor.extract_word_structure(word_path)

            config = MatcherConfig.load()
            h_config = config.get("hierarchy", MatcherConfig.get_defaults()["hierarchy"])
            l1_c = level1_col if level1_col is not None else h_config.get("level1_col", 1)
            l2_c = level2_col if level2_col is not None else h_config.get("level2_col", 2)
            l3_c = level3_col if level3_col is not None else h_config.get("level3_col", 3)

            target_sheet_name = sheet_name
            if target_sheet_name is None:
                sheet_idx = h_config.get("sheet_name", 2)
                try:
                    xl = pd.ExcelFile(excel_path)
                    sheet_names = xl.sheet_names
                    if len(sheet_names) > sheet_idx:
                        target_sheet_name = sheet_names[sheet_idx]
                    elif len(sheet_names) > 0:
                        target_sheet_name = sheet_names[0]
                except Exception as e:
                    log_error(f'Error checking sheets: {e}')
                    target_sheet_name = 0

            if project_name:
                clean_name = ReportGenerator._clean_project_name(project_name)
            elif word_sections_preloaded is not None:
                clean_name = "项目报告"
            else:
                clean_name = ReportGenerator._clean_project_name(word_path)

            # 【核心修复】：统一路径模块生成本次运行的时间戳文件夹
            output_dir = run_output_dir(
                project_name or word_path,
                report_base_name,
            )

            # 文件名统一带上项目名
            tree_log_path = os.path.join(output_dir, f"{clean_name}_Word结构树.txt")

            report = matcher.match_hierarchical_documents(
                word_path, excel_path, sheet_name=target_sheet_name,
                level1_col=l1_c, level2_col=l2_c, level3_col=l3_c, header=header_row,
                progress_callback=progress_callback, hierarchy_log_path=tree_log_path,
                word_sections_preloaded=word_sections_preloaded,
            )

            report_path = os.path.join(output_dir, f"{clean_name}_层级匹配报告.xlsx")
            saved_path = matcher.save_report(report, report_path)
            report["report_path"] = os.path.abspath(saved_path)

            # 保存 JSON
            json_report_dir = os.path.join(output_dir, "json_reports")
            if not os.path.exists(json_report_dir): os.makedirs(json_report_dir, exist_ok=True)
            json_report_path = os.path.join(json_report_dir, "hierarchy_matching.json")
            try:
                import json
                with open(json_report_path, "w", encoding="utf-8") as f:
                    json.dump(report, f, ensure_ascii=False, indent=2)
                report["json_report_path"] = json_report_path
            except Exception as e:
                log_error(f'警告: 无法保存JSON报告: {e}')

            return report
        except Exception as e:
            log_error(f'层级匹配校验失败: {e}')
            return {"is_valid": False, "error": str(e)}

    @staticmethod
    def validate_functional_process(
            word_path, excel_path, header_row=0, func_col=None, sheet_name=None,
            fuzzy_match=True, threshold=0.8, progress_callback=None,
            word_sections_preloaded=None, project_name=None,
            hierarchy_mapping=None, preloaded_word_data=None,
            report_base_name=None,
    ):
        """节点6：功能过程校验"""
        try:
            from docx import Document
            from utils.initial_review_report import report_file_path

            doc = None
            word_file_path = word_path
            if hasattr(word_path, "paragraphs") and hasattr(word_path, "element"):
                doc = word_path
                word_content = word_sections_preloaded or DocumentProcessor.extract_word_structure(doc)
            else:
                word_content = word_sections_preloaded or DocumentProcessor.extract_word_structure(word_path)

            config = MatcherConfig.load()
            p_config = config.get("process", MatcherConfig.get_defaults()["process"])
            h_config = config.get("hierarchy", MatcherConfig.get_defaults()["hierarchy"])
            func_proc_col = func_col if func_col is not None else p_config.get("column", 6)
            l1_col, l2_col, l3_col = h_config.get("level1_col", 1), h_config.get("level2_col", 2), h_config.get(
                "level3_col", 3)

            matcher = HierarchicalMatcher(fuzzy_match=fuzzy_match, threshold=threshold)
            target_sheet = sheet_name
            if target_sheet is None:
                sheet_idx = p_config.get("sheet_name", 2)
                try:
                    xl = pd.ExcelFile(excel_path)
                    sheet_names = xl.sheet_names
                    target_sheet = sheet_names[sheet_idx] if len(sheet_names) > sheet_idx else (
                        sheet_names[0] if sheet_names else 0)
                except:
                    target_sheet = 0

            try:
                df_header = pd.read_excel(excel_path, sheet_name=target_sheet, nrows=0, header=header_row)
                if func_proc_col >= len(df_header.columns): func_proc_col = 0
            except:
                pass

            preloaded_data = None
            if preloaded_word_data:
                excel_data = preloaded_word_data.get("excel_data") or matcher.extract_excel_content(excel_path,
                                                                                                    mode="flat",
                                                                                                    sheet_name=target_sheet,
                                                                                                    header=header_row,
                                                                                                    column=func_proc_col,
                                                                                                    level1_col=l1_col,
                                                                                                    level2_col=l2_col,
                                                                                                    level3_col=l3_col)
                preloaded_data = {**preloaded_word_data, "excel_data": excel_data}

            report = matcher.match_documents(
                word_file_path, excel_path, sheet_name=target_sheet, header=header_row,
                column=func_proc_col, level1_col=l1_col, level2_col=l2_col, level3_col=l3_col,
                hierarchy_mapping=hierarchy_mapping, full_text_search=True,
                progress_callback=progress_callback,
                word_items_preloaded=(word_content if isinstance(word_content, list) else None),
                preloaded_data=preloaded_data,
            )

            # 【核心修复】统一路径模块生成本次运行时间戳文件夹内的报告路径
            _, report_path = report_file_path(
                report_type="功能过程报告",
                project_name_or_path=project_name or word_path,
                report_base_name=report_base_name,
            )

            saved_path = matcher.save_report(report, report_path)
            report["report_path"] = os.path.abspath(saved_path)
            return report
        except Exception as e:
            log_error(f'功能过程校验失败: {e}')
            return {"is_valid": False, "error": str(e)}

    @staticmethod
    def validate_data_movement_types(
            excel_path, sheet_name=None, header_row=0, func_col=6, move_col=8,
            progress_callback=None, project_name=None, report_base_name=None,
    ):
        """节点7：功能过程数据移动类型校验"""
        try:
            import pandas as pd
            from utils.initial_review_report import report_file_path

            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
            if func_col >= len(df.columns) or move_col >= len(df.columns):
                return {"is_valid": False, "error": f"Excel 列索引越界"}

            df_process = df.copy()
            df_process.iloc[:, func_col] = df_process.iloc[:, func_col].ffill()
            results, current_process, current_moves, process_rows = [], None, [], []

            def check_moves(moves):
                if not moves: return "无数据移动"
                valid_moves = [str(m).strip().upper() for m in moves if str(m).strip().upper() in ["E", "R", "W", "X"]]
                if not valid_moves: return "缺少有效类型"
                if valid_moves[0] != "E": return "缺少E"
                if valid_moves[-1] not in ["W", "X"]: return "缺少X"
                return "合规"

            for idx, row in df_process.iterrows():
                proc_val = str(row.iloc[func_col]).strip()
                move_val = str(row.iloc[move_col]).strip()
                if not proc_val or proc_val.lower() == "nan" or "体现了" in proc_val or "功能过程" in proc_val: continue
                if current_process != proc_val:
                    if current_process:
                        results.append({"process": current_process, "moves": "".join(current_moves),
                                        "result": check_moves(current_moves),
                                        "row_range": f"{process_rows[0] + header_row + 2}-{process_rows[-1] + header_row + 2}" if len(
                                            process_rows) > 1 else f"{process_rows[0] + header_row + 2}"})
                    current_process, current_moves, process_rows = proc_val, [], []
                if move_val.upper() in ["E", "R", "W", "X"]: current_moves.append(move_val.upper())
                process_rows.append(idx)

            if current_process:
                results.append(
                    {"process": current_process, "moves": "".join(current_moves), "result": check_moves(current_moves),
                     "row_range": f"{process_rows[0] + header_row + 2}-{process_rows[-1] + header_row + 2}" if len(
                         process_rows) > 1 else f"{process_rows[0] + header_row + 2}"})

            passed = sum(1 for r in results if r["result"] == "合规")
            failed = len(results) - passed
            report_path = None

            if len(results) > 0:
                try:
                    df_res = pd.DataFrame(results).rename(
                        columns={"process": "功能过程", "moves": "移动类型序列", "result": "校验结果",
                                 "row_range": "Excel行号"})

                    # 【核心修复】统一路径模块生成本次运行时间戳文件夹内的报告路径
                    _, report_path = report_file_path(
                        report_type="数据移动类型报告",
                        project_name_or_path=project_name or excel_path,
                        report_base_name=report_base_name,
                    )

                    # 防重名处理
                    counter = 1
                    base_report_path = report_path
                    while os.path.exists(report_path):
                        name, ext = os.path.splitext(base_report_path)
                        report_path = f"{name}({counter}){ext}"
                        counter += 1

                    df_res.to_excel(report_path, index=False)
                    report_path = os.path.abspath(report_path)
                except Exception as e:
                    log_error(f'Error saving data movement report: {e}')

            return {"is_valid": failed == 0, "items": results,
                    "statistics": {"总数": len(results), "合规": passed, "不合规": failed}, "report_path": report_path}
        except Exception as e:
            import traceback
            log_error(f'数据移动类型校验失败: {e}')
            log_error(traceback.format_exc())
            return {"is_valid": False, "error": str(e)}
