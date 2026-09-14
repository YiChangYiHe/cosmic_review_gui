from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QWidget,
    QMenu, QApplication, QSizePolicy, QStackedWidget, QTextBrowser,
    QTableWidget, QTableWidgetItem, QHeaderView, QButtonGroup,
    QScrollArea, QGridLayout,
)
from PySide6.QtGui import QAction, QColor, QFont, QIcon
from PySide6.QtCore import Qt, QThread, Signal, QTimer
import time
import os
import json
from .steps_widget import StepsWidget
from .report_dialog import ReportDialog, SummaryDialog
from utils.document_processor import DocumentProcessor
from utils.similarity_checker import SimilarityChecker
from utils.report_generator import ReportGenerator
from utils.asset_matcher import compare_split_with_asset
from utils.data_attribute_checker import check_data_attribute_duplicates
from utils.path_utils import get_resource_path, clean_project_name, open_directory
from utils.initial_review_report import (
    make_run_base_name,
    run_output_dir,
    latest_run_dir,
)
from utils.runtime_logger import RuntimeLogger
from extend.matcher_config import MatcherConfig
from extend.hierarchical_matcher import HierarchicalMatcher
from utils.runtime_logger import log_debug, log_error, log_info, log_warn
from utils.task_state import begin_heavy_task_if_large, end_heavy_task


class ElidedTitleLabel(QLabel):
    """[NEW] 超长标题省略号显示：宽度自适应容器，不撑大整个界面"""

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self._full_text = text
        super().setText(text)
        self.setToolTip(text)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.setMinimumWidth(80)

    def setText(self, text):
        self._full_text = text
        self.setToolTip(text)
        super().setText(self._elide())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        super().setText(self._elide())

    def _elide(self):
        if self.width() <= 20:
            return self._full_text
        fm = self.fontMetrics()
        return fm.elidedText(self._full_text, Qt.ElideRight, max(50, self.width() - 4))


def build_tree_text(items):
    """由标题项列表构建目录树文本"""
    lines = ["📋 Word 文档完整目录树结构", "=" * 60]
    for item in items or []:
        if not isinstance(item, dict):
            continue
        title = item.get("text") or item.get("title") or ""
        num = str(item.get("number") or "")
        if not title:
            continue
        depth = len([p for p in num.split(".") if p.strip()]) or 1
        indent = "   " * max(0, depth - 1)
        lines.append(f"{indent}[{num}] {title}")
    lines.append("=" * 60)
    return "\n".join(lines)


class ValidationWorker(QThread):
    """后台校验线程，支持进度反馈"""
    progress = Signal(int, str)  # 进度值, 子步骤描述
    step_result = Signal(int, dict)  # step_num, result_data
    finished = Signal(dict)

    def __init__(self, task_data):
        super().__init__()
        self.task_data = task_data
        self._is_running = True
        self._finished_normally = False  # [NEW] 标记任务是否正常完成
        self._current_v_res = {}  # [NEW] 用于实时保存进度

    def _save_run_info(self, raw_info):
        try:
            import json as _json
            # 【修复】使用统一路径模块获取带时间戳的目录
            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            # 文件名固定为 run_info.json
            path = os.path.join(out_dir, "run_info.json")

            temp_path = path + ".tmp"
            with open(temp_path, "w", encoding="utf-8") as f:
                _json.dump(raw_info or {}, f, ensure_ascii=False, default=str)
                f.flush();
                os.fsync(f.fileno())
            os.replace(temp_path, path)
        except Exception as e:
            log_warn(f'[WARN] 保存运行信息失败: {e}')
    def _clear_pause_state(self):
        try:
            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            path = os.path.join(out_dir, "paused.json")
            if os.path.exists(path): os.remove(path)
        except:
            pass

    def _save_tree_txt(self, tree_text):
        try:
            # 结构树报告也放入时间戳文件夹
            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            path = os.path.join(out_dir, "Word结构树.txt")

            with open(path, "w", encoding="utf-8") as f:
                f.write(tree_text)
            self._tree_txt_path = path
            RuntimeLogger.log(f"✅ [Step 0] 结构树txt已保存: {path}")
        except Exception as e:
            RuntimeLogger.log(f"⚠️ [Step 0] 结构树txt保存失败: {e}", level="WARN")

    _STEP_KEY_MAP = [
        (9, "data_attr_res"),
        (8, "asset_res"),
        (7, "move_res"),
        (6, "process_res"),
        (5, "hierarchy_res"),
        (4, "factor_check"),
        (3, "ratio_check"),
        (2, "excel_check"),
        (1, "suspect_modules"),
    ]

    def _dump_pause_state(self, v_res=None):
        try:
            import json as _json
            v_res = v_res or getattr(self, "_current_v_res", {}) or {}
            completed = -1
            for s, k in self._STEP_KEY_MAP:
                if k in v_res: completed = max(completed, s)
            if completed < 1: return

            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            path = os.path.join(out_dir, "paused.json")

            with open(path, "w", encoding="utf-8") as f:
                _json.dump({"raw_task_info": self.task_data.get("raw_task_info", {}), "v_res": v_res,
                            "completed_through": completed}, f, ensure_ascii=False, default=str)
            RuntimeLogger.log(f"⏸ 任务已暂停，进度已保存（已完成前 {completed} 步）")
        except Exception as e:
            log_warn(f'[WARN] 保存暂停进度失败: {e}')
    def _save_result_json(self, v_res):
        """保存完整运行结果，供历史页面分析状态"""
        try:
            import json
            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            path = os.path.join(out_dir, "result.json")
            # 保存 v_res 和完成标记 (10表示全部10个节点完成)
            data = {"v_res": v_res, "completed_through": 10}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, default=str)
        except Exception as e:
            log_warn(f'[WARN] 保存 result.json 失败: {e}')

    def _save_summary_html(self, v_res):
        """任务完成后生成结果汇总 HTML 并记录到 run_info.json，供历史页面查看"""
        try:
            from ui.report_dialog import generate_summary_html
            out_dir = run_output_dir(
                self.task_data.get("filename", ""),
                getattr(self, "_run_base_name", ""),
            )
            html = generate_summary_html({
                "filename": self.task_data.get("filename", ""),
                "validation_results": [v_res],
            })
            path = os.path.join(out_dir, "summary.html")
            temp_path = path + ".tmp"
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(html)
                f.flush()
                os.fsync(f.fileno())
            os.replace(temp_path, path)

            run_info_path = os.path.join(out_dir, "run_info.json")
            run_data = {}
            if os.path.exists(run_info_path):
                try:
                    with open(run_info_path, "r", encoding="utf-8") as f:
                        run_data = json.load(f) or {}
                except Exception:
                    run_data = {}
            run_data["summary_html_path"] = os.path.abspath(path)
            temp_info = run_info_path + ".tmp"
            with open(temp_info, "w", encoding="utf-8") as f:
                json.dump(run_data, f, ensure_ascii=False, default=str)
                f.flush()
                os.fsync(f.fileno())
            os.replace(temp_info, run_info_path)
            RuntimeLogger.log(f"✅ 结果汇总已保存: {path}")
        except Exception as e:
            log_warn(f'[WARN] 保存结果汇总失败: {e}')

    def stop(self):
        """停止线程"""
        self._is_running = False

    def run(self):
        # ================= 【新增】读取并应用用户设置的日志级别 =================
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        log_level = config.get("log_level", "INFO")
        RuntimeLogger.set_log_level(log_level)
        # ======================================================================

        # 设置项目名称用于日志
        project_name = self.task_data.get("filename", "UnknownProject")
        RuntimeLogger.set_project(project_name)

        # 【新增】大项目检测：输入文件总量超阈值则提升线程优先级并暂缓历史页扫描
        _early_raw = self.task_data.get("raw_task_info", {}) or {}
        _heavy_paths = []
        for _pair in (_early_raw.get("file_pairs", []) or []):
            if isinstance(_pair, dict):
                for _k in ("word", "excel", "asset"):
                    _heavy_paths.append(_pair.get(_k))
        self._heavy_task = begin_heavy_task_if_large(_heavy_paths, project_name, kind="initial")

        # [NEW] 本次运行统一的基础名（初评产物路径统一由 initial_review_report 模块生成）
        self._run_base_name = make_run_base_name(project_name)
        self._save_run_info(self.task_data.get("raw_task_info", {}) or {})

        # [NEW] 断点续跑：从暂停状态恢复
        resume_state = (self.task_data.get("raw_task_info", {}) or {}).get(
            "resume_state"
        ) or {}
        try:
            self._resumed_v_res = dict(resume_state.get("v_res", {}) or {})
            self._completed_through = int(resume_state.get("completed_through", -1))
        except Exception:
            self._resumed_v_res = {}
            self._completed_through = -1

        self._resume_source_paused = (self.task_data.get("raw_task_info", {}) or {}).get(
            "resume_source_paused_path"
        )

        if self._completed_through >= 1:
            RuntimeLogger.log(
                f"▶ 检测到暂停进度：已完成前 {self._completed_through} 步，将从下一步继续"
            )
            for _s in range(0, self._completed_through + 1):
                self.step_result.emit(_s, self._resumed_v_res)

        RuntimeLogger.log(f"开始后台校验任务...")

        try:
            # 【并行安全】只清理历史运行遗留的陈旧临时文件，不再全量清空
            # （全量清理会删除其他并行任务正在使用的缓存与临时文件导致崩溃）
            DocumentProcessor.cleanup_stale_temp()
            self.progress.emit(1, "正在初始化校验环境...")

            raw_info = self.task_data.get("raw_task_info", {})
            file_pairs = raw_info.get("file_pairs", [])
            RuntimeLogger.log(f"获取到文件配对数量: {len(file_pairs)}")

            if not file_pairs:
                RuntimeLogger.log(
                    "️ 错误: file_pairs 列表为空，无法继续。", level="ERROR"
                )
                self.finished.emit({})
                return

            # ================= 【新增】读取自动编号配置 =================
            auto_numbering = raw_info.get("auto_numbering", False)
            # ============================================================

            # 1. 启动并行加载架构
            import concurrent.futures
            pair = file_pairs[0]
            template_path = get_resource_path("folder/附件1XX项目需求说明书V1.0.0.docx")

            RuntimeLogger.log(
                f" [Step 0] 环境准备 - 启动并行架构，Word文档处理前置到此阶段完成..."
            )
            self.progress.emit(2, "正在并发加载 Word/Excel 文件...")

            with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
                # 定义核心预加载任务
                futures = {}
                if template_path and os.path.exists(template_path):
                    futures['tpl'] = executor.submit(
                        DocumentProcessor.load_word_document, template_path
                    )
                if pair.get("word") and os.path.exists(pair["word"]):
                    futures['target'] = executor.submit(
                        DocumentProcessor.load_word_document, pair["word"]
                    )
                futures['wb'] = executor.submit(
                    DocumentProcessor.load_excel_workbook, pair["excel"], True
                )
                futures['excel_info'] = executor.submit(
                    DocumentProcessor.extract_excel_info, pair["excel"]
                )

                # 等待加载基础对象
                tpl_doc = futures.get('tpl', None)
                if tpl_doc:
                    tpl_doc = tpl_doc.result()
                target_doc = futures.get('target', None)
                if target_doc:
                    target_doc = target_doc.result()
                target_wb = futures['wb'].result()

                self.progress.emit(8, "正在提取文档深层结构 (目录/层级)...")

                # === 新增：数据预处理阶段 ===
                run_simple = raw_info.get("run_simple", False)
                word_preprocessing_success = False
                step0_emitted = False

                if run_simple and pair["word"] and pair["excel"]:
                    RuntimeLogger.log(
                        f"🔄 [Step 0.5] 开始预处理Word和Excel数据以优化后续匹配..."
                    )
                    self.progress.emit(10, "Word文档打开与大纲提取...")

                    try:
                        def word_prep_progress(p, msg):
                            mapped_p = 10 + int(p * 0.12)
                            self.progress.emit(mapped_p, f"Word预处理: {msg}")

                        RuntimeLogger.log(
                            f"  [INFO]  [Step 0] 开始完整Word大纲及内容预处理..."
                        )

                        if pair["word"] and os.path.exists(pair["word"]):
                            try:
                                matcher = HierarchicalMatcher()
                                word_prep_progress(10, "正在启动 Word 结构及全文预提取...")

                                tree_txt = raw_info.get("tree_txt")
                                word_cache = None

                                if tree_txt and os.path.exists(tree_txt):
                                    RuntimeLogger.log(f"  [INFO] 使用导入的目录树txt: {tree_txt}")
                                    word_prep_progress(20, "正在解析导入的目录树txt...")
                                    word_cache = matcher.build_cache_from_tree_txt(
                                        tree_txt, pair["word"],
                                        auto_numbering=auto_numbering,
                                        project_name=self.task_data.get("filename"),
                                    )
                                else:
                                    word_cache = matcher.prepare_word_data_async(
                                        pair["word"],
                                        auto_numbering=auto_numbering,
                                        project_name=self.task_data.get("filename"),
                                    )

                                if word_cache and word_cache.get("items"):
                                    word_items = word_cache["items"]
                                    full_text_content = word_cache.get("full_text_content", [])

                                    word_prep_progress(60, f"已提取 {len(word_items)} 个章节，正在集成数据...")
                                    word_prep_progress(70, "正在预加载 Excel 功能点数据...")

                                    try:
                                        excel_config = raw_info
                                        excel_sheet = excel_config.get("simple_sheet", 2)
                                        excel_header = excel_config.get("functional_header_row", 0)
                                        excel_col = excel_config.get("functional_column_index", 6)

                                        excel_data = matcher.extract_excel_content(
                                            pair["excel"], mode="flat",
                                            sheet_name=excel_sheet, header=excel_header,
                                            column=excel_col, level1_col=1,
                                            level2_col=2, level3_col=3,
                                        )
                                        word_cache["excel_data"] = excel_data
                                        RuntimeLogger.log(
                                            f"  [INFO] Excel内容预加载完成: {len(excel_data) if excel_data else 0} 项")
                                    except Exception as e:
                                        RuntimeLogger.log(f"  [WARN] Excel预加载失败: {str(e)[:80]}", level="WARN")
                                        word_cache["excel_data"] = None

                                    if not hasattr(self, "data_cache"):
                                        self.data_cache = {}
                                    self.data_cache["word_data"] = word_cache

                                    word_prep_progress(100, f"完成: {len(word_items)} 项已缓存，包含全文内容")
                                    RuntimeLogger.log(
                                        f"✅ [Step 0] Word完整数据就绪: {len(word_items)} 章节, 全文 {len(full_text_content)} 项")

                                    word_preprocessing_success = True

                                    if not step0_emitted:
                                        step0_emitted = True
                                        tree_text = build_tree_text(word_items)
                                        self._save_tree_txt(tree_text)
                                        self.step_result.emit(
                                            0, {
                                                "env_check": {"is_ok": True, "sections": len(word_items)},
                                                "tree_text": tree_text,
                                            }
                                        )
                                else:
                                    RuntimeLogger.log(f"⚠️ Word预处理返回空结果", level="WARN")
                            except Exception as e:
                                RuntimeLogger.log(f"️ Word数据预处理异常: {str(e)[:150]}", level="ERROR")
                        else:
                            RuntimeLogger.log(f"⚠️ Word文件不存在", level="WARN")
                    except Exception as e:
                        RuntimeLogger.log(f"⚠️ Word数据预处理异常: {e}")
                        RuntimeLogger.log(f"  [INFO] 将使用标准Word处理流程")

                self.progress.emit(22, "正在继续提取文档结构...")

                # 开始并发提取详细结构
                run_template = raw_info.get("run_template", True)
                run_factors = raw_info.get("run_factors", True)
                run_hierarchy = raw_info.get("run_hierarchy", True)

                need_target_structure = any([run_template, run_factors, run_hierarchy, run_simple])
                need_tpl_structure = run_template

                template_sections = []
                target_sections = []

                RuntimeLogger.log(
                    f"[DEBUG] word_preprocessing_success={word_preprocessing_success}, need_target_structure={need_target_structure}"
                )

                if need_target_structure and not word_preprocessing_success:
                    RuntimeLogger.log(f" [Step 0] 正在预提取 Word 目录树... (优先使用稳定大纲提取)")
                    matcher = HierarchicalMatcher()

                    def extract_word_hierarchy(path, doc_obj):
                        def extraction_progress_proxy(p, msg):
                            mapped_p = 8 + int(p * 0.1)
                            self.progress.emit(mapped_p, msg)

                        tree_txt = raw_info.get("tree_txt")
                        if tree_txt and os.path.exists(tree_txt) and path == pair.get("word"):
                            try:
                                RuntimeLogger.log(f"  [INFO] 使用导入的目录树txt提取结构: {tree_txt}")
                                h_res = HierarchicalMatcher.hierarchy_from_tree_txt(tree_txt)
                                items = h_res.get("all_items", [])
                                if items:
                                    extraction_progress_proxy(100, "目录树txt解析完成")
                                    return items
                            except Exception as e:
                                RuntimeLogger.log(f"  [WARN] 目录树txt解析失败，回退文档解析: {e}", level="WARN")

                        try:
                            if path and os.path.exists(path):
                                RuntimeLogger.log(
                                    f"  [INFO]  使用 HierarchicalMatcher 提取标题结构: {os.path.basename(path)}")
                                extraction_progress_proxy(0, "正在通过 Win32 接口提取 Word 标题结构...")
                                matcher = HierarchicalMatcher()
                                result = matcher.extract_word_outline_as_hierarchy(
                                    path, include_content=False, auto_numbering=auto_numbering
                                )
                                if result:
                                    items = result if isinstance(result, list) else result.get("all_items", [])
                                    extraction_progress_proxy(100, "Word 结构提取完成")
                                    RuntimeLogger.log(f"  [OK] ✅ 标题结构提取成功: {len(items)} 项")
                                    return items

                            RuntimeLogger.log(
                                f"  [WARN] ️ 降级使用 DocumentProcessor.extract_word_structure(use_stable=True)...")
                            result = DocumentProcessor.extract_word_structure(
                                path, use_stable=True, progress_callback=extraction_progress_proxy,
                            )
                            if result and len(result) > 0:
                                RuntimeLogger.log(f"  [OK] DocumentProcessor 提取成功: {len(result)} 项")
                                return result

                            RuntimeLogger.log(
                                f"  [WARN] ️ 降级使用 HierarchicalMatcher.extract_outline_from_docx_object...")
                            result = matcher.extract_outline_from_docx_object(doc_obj)
                            if result.get("all_items") and len(result["all_items"]) > 0:
                                RuntimeLogger.log(f"  [OK] HierarchicalMatcher 提取成功: {len(result['all_items'])} 项")
                                return result["all_items"]

                            raise Exception("无法通过任何方式提取文档结构")
                        except Exception as e:
                            import traceback
                            RuntimeLogger.log(f"[ERROR] extract_word_hierarchy 异常: {e}", level="ERROR")
                            RuntimeLogger.log(traceback.format_exc(), level="DEBUG")

                            RuntimeLogger.log(
                                f"[WARN] 🛟 最终降级: DocumentProcessor.extract_word_structure(无 use_stable)...")
                            try:
                                fallback = DocumentProcessor.extract_word_structure(doc_obj)
                                RuntimeLogger.log(f"  [OK] 🛟 最终降级成功: {len(fallback) if fallback else 0} 项")
                                return fallback or []
                            except Exception as e2:
                                RuntimeLogger.log(f"[ERROR] 所有提取方案均失败: {e2}", level="ERROR")
                                return []

                    structure_futures = {}
                    if need_tpl_structure:
                        structure_futures["tpl"] = executor.submit(
                            extract_word_hierarchy, template_path, tpl_doc
                        )
                        if pair.get("word") and os.path.exists(pair["word"]):
                            structure_futures["target"] = executor.submit(
                                extract_word_hierarchy, pair["word"], target_doc
                            )

                        excel_info = futures['excel_info'].result()

                        if "tpl" in structure_futures:
                            template_sections = structure_futures["tpl"].result()
                        if "target" in structure_futures:
                            target_sections = structure_futures["target"].result()

                            if target_sections and not step0_emitted:
                                step0_emitted = True
                                tree_text = build_tree_text(target_sections)
                                self._save_tree_txt(tree_text)
                                self.step_result.emit(
                                    0, {
                                        "env_check": {"is_ok": True, "sections": len(target_sections)},
                                        "tree_text": tree_text,
                                    }
                                )

                    RuntimeLogger.log(f"✅ [Step 0] 预提取完成，共获取 {len(target_sections)} 个章节/内容项")

                elif word_preprocessing_success:
                    RuntimeLogger.log(" [Step 0] 使用预处理的Word数据，跳过重复提取")
                    excel_info = futures['excel_info'].result()
                    target_sections = getattr(self, "data_cache", {}).get("word_data", {}).get("items", [])

                    if need_tpl_structure and not template_sections:
                        RuntimeLogger.log("🔎 [Step 0] 预处理模式：同步加载模板文档结构 (含正文)...")
                        try:
                            matcher = HierarchicalMatcher()
                            tpl_res = matcher.extract_word_outline_as_hierarchy(
                                template_path, include_content=True, auto_numbering=auto_numbering
                            )
                            template_sections = tpl_res.get("all_items", [])
                            RuntimeLogger.log(f"✅ [Step 0] 模板结构加载完成: {len(template_sections)} 项")
                        except Exception as e:
                            RuntimeLogger.log(f"⚠️ [Step 0] 模板结构加载失败: {e}", level="WARN")

                    RuntimeLogger.log(f"✅ [Step 0] 使用缓存数据: {len(target_sections)} 个章节/内容项")
                else:
                    RuntimeLogger.log(" [Step 0] 跳过 Word 结构提取 (未选择任何 Word 校验节点)")
                    excel_info = futures['excel_info'].result()

            if not tpl_doc and need_tpl_structure:
                RuntimeLogger.log("⚠️ 模板文件加载失败", level="WARN")

            has_word = pair.get("word") is not None and os.path.exists(pair["word"])
            if not target_doc and need_target_structure and has_word:
                RuntimeLogger.log("❌ 目标 Word 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Word 文件"})
                return

            if not target_wb:
                RuntimeLogger.log("❌ 目标 Excel 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Excel 文件"})
                return

            self.progress.emit(15, "正在扫描 Excel 工作表列表...")
            RuntimeLogger.log(f"✅ [Step 0] 基础对象并发解析完成 (目标项: {len(target_sections)})")

            if not self._is_running:
                self._dump_pause_state()
                return

            if not step0_emitted:
                self.step_result.emit(0, {"env_check": {"is_ok": True, "sections": len(target_sections)}})

            # [NEW] 断点续跑：沿用上次已完成步骤的结果
            v_res = dict(getattr(self, "_resumed_v_res", {}) or {})
            self._current_v_res = v_res

            self.progress.emit(15, "正在扫描 Excel 工作表列表...")
            check_sheet = (
                raw_info.get("hierarchy_sheet")
                if raw_info.get("run_hierarchy")
                else raw_info.get("simple_sheet")
            )

            # 4. 节点 1：模板合规性校验
            step1_start = time.time()
            if "suspect_modules" in self._resumed_v_res:
                pass
            elif raw_info.get("run_template", True):
                has_word = pair.get("word") and os.path.exists(pair["word"])
                if has_word and target_sections:
                    RuntimeLogger.log(f"正在启动 [Step 1] 模板合规性比对校验...")
                    self.progress.emit(20, "正在比对 Word 章节与标准模板...")
                    res_s1 = SimilarityChecker.validate_template(
                        target_sections, excel_info, template_sections
                    )
                    v_res.update(res_s1)
                    self._current_v_res = v_res  # [NEW] 立即保存
                else:
                    RuntimeLogger.log(f"[Step 1] 跳过模板校验：未上传Word文档")
                    v_res.update({"is_valid": True, "skipped": True, "reason": "未上传Word文档"})
            else:
                v_res.update({"is_valid": True, "skipped": True})

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["duration"] = time.time() - step1_start
            self.step_result.emit(1, v_res)

            # [UI 对齐] 立即跃迁至第 2 步起始进度 (30%)
            self.progress.emit(30, "正在准备 Excel 空值扫描...")

            # 5. 节点 2：Excel 空值校验
            step2_start = time.time()
            if "excel_check" in self._resumed_v_res:
                pass
            elif raw_info.get("run_empty", True):
                RuntimeLogger.log(f"正在启动 [Step 2] Excel 关键列空值扫描 (Sheet: {check_sheet})...")
                self.progress.emit(31, f"正在扫描 Excel ({check_sheet}) 空值行...")
                excel_check_res = DocumentProcessor.check_excel_empty_cells(
                    pair["excel"], sheet_name=check_sheet,
                    report_base_name=getattr(self, "_run_base_name", None),
                )
                v_res["excel_check"] = excel_check_res
                self._current_v_res = v_res  # [NEW] 立即保存
            else:
                v_res["excel_check"] = {"is_ok": True, "skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["excel_check"]["duration"] = time.time() - step2_start
            self.step_result.emit(2, {"excel_check": v_res["excel_check"]})

            # [UI 对齐] 立即跃迁至第 3 步起始进度 (33%)
            self.progress.emit(33, "正在准备送审比例计算...")

            # 6. 功能匹配校验 (辅助数据)
            if target_sections and (raw_info.get("run_hierarchy") or raw_info.get("run_simple")):
                RuntimeLogger.log(f"正在进行 [辅助步骤] Word 与 Excel 模块名称匹配度计算...")
                excel_modules = DocumentProcessor.get_excel_modules(target_wb, sheet_name=check_sheet)
                func_match = SimilarityChecker.validate_function_matching(target_sections, excel_modules)
                v_res["func_match"] = func_match
            else:
                v_res["func_match"] = {"skipped": True}

            # 7. 节点 3：送审比例校验
            step3_start = time.time()
            if "ratio_check" in self._resumed_v_res:
                pass
            elif raw_info.get("run_ratio", True):
                RuntimeLogger.log(f"正在进行 [Step 3] 送审比例计算...")
                self.progress.emit(34, "当前正在计算送审功能点与人天比例...")
                fp_count = DocumentProcessor.get_functional_points_count(target_wb, sheet_name=check_sheet)
                mandays = float(self.task_data.get("days", 0))
                ratio = fp_count / mandays if mandays > 0 else 0
                r_range = "0.8 ~ 2.0" if mandays <= 1000 else "0.8 ~ 1.5"
                v_res["ratio_check"] = {
                    "fp_count": fp_count, "mandays": mandays, "ratio": round(ratio, 2),
                    "is_ok": ((0.8 <= ratio <= 2.0) if mandays <= 1000 else (0.8 <= ratio <= 1.5)),
                    "range": r_range,
                }
                self._current_v_res = v_res  # [NEW] 立即保存
            else:
                v_res["ratio_check"] = {"is_ok": True, "skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["ratio_check"]["duration"] = time.time() - step3_start
            self.step_result.emit(3, {"ratio_check": v_res["ratio_check"]})

            # [UI 对齐] 立即跃迁至第 4 步起始进度 (36%)
            self.progress.emit(36, "正在准备附加值调整因子提取...")

            # 8. 节点 4：附加值调整因子校验
            step4_start = time.time()
            if "factor_check" in self._resumed_v_res:
                pass
            elif raw_info.get("run_factors", True):
                has_word = pair.get("word") and os.path.exists(pair["word"])
                if has_word and target_doc:
                    RuntimeLogger.log(f"正在进行 [Step 4] Word 附加值调整因子提取...")
                    self.progress.emit(37, "正在扫描文档中的因子表与描述文字...")
                    factors = DocumentProcessor.check_adjustment_factors_in_word(
                        target_doc, target_sections=target_sections
                    )
                    v_res["factor_check"] = factors
                    self._current_v_res = v_res  # [NEW] 立即保存
                else:
                    RuntimeLogger.log(f"[Step 4] 跳过附加值因子：未上传Word文档")
                    v_res["factor_check"] = {"skipped": True, "reason": "未上传Word文档"}
            else:
                v_res["factor_check"] = {"skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["factor_check"]["duration"] = time.time() - step4_start
            self.step_result.emit(4, {"factor_check": v_res["factor_check"]})

            # [UI 对齐] 立即跃迁至第 5 步起始进度 (39%)
            self.progress.emit(39, "正在启动核心层级匹配引擎...")

            # 9. 层级匹配校验 (Node 5)
            step5_start = time.time()
            run_hierarchy = raw_info.get("run_hierarchy", True)
            fuzzy = raw_info.get("fuzzy", True)
            threshold = raw_info.get("threshold", 0.8)
            has_word = pair.get("word") is not None and os.path.exists(pair["word"])
            hierarchy_mapping = None

            if "hierarchy_res" in self._resumed_v_res:
                pass
            elif run_hierarchy and has_word:
                RuntimeLogger.log(f"正在启动 [Step 5] 核心层级匹配校验...")

                def hierarchy_progress_proxy(p, msg):
                    if not self._is_running:
                        self._dump_pause_state()
                        return
                    mapped_progress = 39 + int(p * 0.31)
                    self.progress.emit(mapped_progress, f"层级匹配: {msg}")
                    if msg and (
                            "开始" in msg or "完成" in msg or "1/" in msg or "/100" in msg or "[PROCESS]" in msg or "阶段" in msg or "匹配" in msg):
                        log_debug(f"[Step 5] {msg}")

                h_header_row = raw_info.get("hierarchy_header_row", 0)
                l1_col = raw_info.get("level1_column_index", 1)
                l2_col = raw_info.get("level2_column_index", 2)
                l3_col = raw_info.get("level3_column_index", 3)
                hier_sheet = raw_info.get("hierarchy_sheet")

                hierarchy_res = DocumentProcessor.validate_hierarchy_matching(
                    pair["word"], pair["excel"],
                    header_row=h_header_row - 1, level1_col=l1_col,
                    level2_col=l2_col, level3_col=l3_col, sheet_name=hier_sheet,
                    fuzzy_match=fuzzy, threshold=threshold,
                    progress_callback=hierarchy_progress_proxy,
                    word_sections_preloaded=target_sections,
                    project_name=self.task_data["filename"],
                    report_base_name=getattr(self, "_run_base_name", None),
                )
                v_res["hierarchy_res"] = hierarchy_res
                self._current_v_res = v_res  # [NEW] 立即保存

                try:
                    hierarchy_mapping = {}
                    matched_items = hierarchy_res.get("exact_matched", []) + hierarchy_res.get("fuzzy_matched", [])
                    for item in matched_items:
                        e_l1 = str(item.get("Excel一级模块", "")).strip()
                        e_l2 = str(item.get("Excel二级模块", "")).strip()
                        e_l3 = str(item.get("Excel三级模块", "")).strip()
                        key = (e_l1, e_l2, e_l3)
                        w_title = item.get("Word三级标题") or item.get("Word二级标题") or item.get("Word一级标题")
                        if w_title:
                            hierarchy_mapping[key] = w_title
                    RuntimeLogger.log(f"已成功提取 {len(hierarchy_mapping)} 个层级映射锚点用于加速后续匹配")
                except Exception as e:
                    RuntimeLogger.log(f"提取层级映射失败: {e}", level="WARN")

                RuntimeLogger.log(f"层级匹配完成: {hierarchy_res.get('statistics', {})}")
            else:
                v_res["hierarchy_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["hierarchy_res"]["duration"] = time.time() - step5_start
            self.step_result.emit(5, {"hierarchy_res": v_res["hierarchy_res"]})

            # [UI 对齐] 完成第 5 步后，立即将 UI 文字推进至第 6 步区位 (70%)
            self.progress.emit(70, "正在启动功能过程内容匹配...")

            # 10. 功能过程校验 (Node 6)
            step6_start = time.time()
            run_simple = raw_info.get("run_simple", False)
            simple_sheet = raw_info.get("simple_sheet")

            if "process_res" in self._resumed_v_res:
                pass
            elif run_simple and has_word:
                RuntimeLogger.log(f"正在进行功能过程匹配校验 (Sheet: {simple_sheet})...")
                f_header_row = raw_info.get("functional_header_row", 0)
                f_col = raw_info.get("functional_column_index", 6)

                def process_progress_proxy(p, msg):
                    if not self._is_running:
                        self._dump_pause_state()
                        return
                    mapped_progress = 70 + int(p * 0.25)
                    self.progress.emit(mapped_progress, f"过程匹配: {msg}")
                    if msg and (
                            "开始" in msg or "完成" in msg or "1/" in msg or "/100" in msg or "[PROCESS]" in msg or "阶段" in msg or "搜索" in msg or "进度" in msg):
                        log_debug(f"[Step 6] {msg}")

                process_res = DocumentProcessor.validate_functional_process(
                    target_doc, pair["excel"],
                    header_row=f_header_row - 1, func_col=f_col,
                    sheet_name=simple_sheet, fuzzy_match=fuzzy, threshold=threshold,
                    progress_callback=process_progress_proxy,
                    word_sections_preloaded=target_sections,
                    project_name=self.task_data["filename"],
                    hierarchy_mapping=hierarchy_mapping,
                    preloaded_word_data=getattr(self, "data_cache", {}).get("word_data"),
                    report_base_name=getattr(self, "_run_base_name", None),
                )
                v_res["process_res"] = process_res
            else:
                v_res["process_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["process_res"]["duration"] = time.time() - step6_start
            self.step_result.emit(6, {"process_res": v_res["process_res"]})

            # [UI 对齐] 完成第 6 步后，立即将 UI 文字推进至第 7 步 (95%)
            self.progress.emit(95, "正在启动数据移动类型校验...")

            # 11. 功能过程数据移动类型校验 (Node 7)
            step7_start = time.time()
            run_move = raw_info.get("run_move", True)

            if "move_res" in self._resumed_v_res:
                pass
            elif run_move and has_word:
                RuntimeLogger.log(f"正在进行 [Step 7] 数据移动类型合规性校验...")
                f_header_row = raw_info.get("functional_header_row", 0)
                f_col = raw_info.get("functional_column_index", 6)
                move_res = DocumentProcessor.validate_data_movement_types(
                    pair["excel"], sheet_name=simple_sheet,
                    header_row=f_header_row - 1, func_col=f_col,
                    move_col=f_col + 2, project_name=self.task_data["filename"],
                    report_base_name=getattr(self, "_run_base_name", None),
                )
                v_res["move_res"] = move_res
                self._current_v_res = v_res  # [NEW] 立即保存
            else:
                v_res["move_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["move_res"]["duration"] = time.time() - step7_start
            self.step_result.emit(7, {"move_res": v_res["move_res"]})

            # 11.5 资产清单匹配校验 (Node 8)
            step8_start = time.time()
            run_asset = raw_info.get("run_asset", True)
            asset_path = pair.get("asset") or raw_info.get("asset_excel")

            if "asset_res" in self._resumed_v_res:
                pass
            elif run_asset and asset_path and os.path.exists(asset_path):
                RuntimeLogger.log(f"正在进行 [Step 8] 资产清单匹配校验...")
                self.progress.emit(96, "正在启动资产清单匹配...")

                def asset_progress_proxy(p, msg):
                    if not self._is_running:
                        self._dump_pause_state()
                        return
                    mapped_progress = 96 + int(p * 2)
                    self.progress.emit(min(mapped_progress, 98), f"资产匹配: {msg}")

                split_sheet = raw_info.get("hierarchy_sheet")
                split_header_row = raw_info.get("hierarchy_header_row")
                split_col_idx = [
                    raw_info.get("level1_column_index"),
                    raw_info.get("level2_column_index"),
                    raw_info.get("level3_column_index"),
                ]
                asset_sheet = raw_info.get("asset_sheet")
                asset_header_row = raw_info.get("asset_header_row")
                asset_col_idx = [
                    raw_info.get("asset_level1_index"),
                    raw_info.get("asset_level2_index"),
                    raw_info.get("asset_level3_index"),
                ]

                RuntimeLogger.log(
                    f"[Step 8] 使用配置 - 拆分表: sheet={split_sheet}, header_row={split_header_row}, cols={split_col_idx}")
                RuntimeLogger.log(
                    f"[Step 8] 使用配置 - 资产清单: sheet={asset_sheet}, header_row={asset_header_row}, cols={asset_col_idx}")

                asset_res = compare_split_with_asset(
                    split_path=pair["excel"], asset_path=asset_path,
                    split_sheet=split_sheet, asset_sheet=asset_sheet,
                    split_header_row=split_header_row, asset_header_row=asset_header_row,
                    split_col_idx=split_col_idx, asset_col_idx=asset_col_idx,
                    threshold=threshold, progress_callback=asset_progress_proxy,
                    should_cancel=lambda: not self._is_running,
                    # 【关键修复】必须传入本次运行的基础名，否则资产清单匹配报告会
                    # 脱离本次运行的时间戳文件夹，被写到以拆分表文件名命名的新文件夹里
                    report_base_name=getattr(self, "_run_base_name", None),
                    log_callback=lambda msg: log_debug(f"[Step 8] {msg}")
                    if msg and ("表头" in str(msg) or "匹配完成" in str(msg) or "统计" in str(msg) or "报告" in str(
                        msg) or "异常" in str(msg))
                    else None,
                )
                if asset_res is None:
                    return
                v_res["asset_res"] = asset_res
                RuntimeLogger.log(f"资产清单匹配完成: {asset_res.get('statistics', {})}")
            else:
                if not run_asset:
                    reason = "未勾选资产清单匹配节点"
                elif not asset_path:
                    reason = "未检测到资产清单文件（请上传文件名包含资产清单的 Excel 文件)"
                else:
                    reason = "资产清单文件不存在"
                v_res["asset_res"] = {"is_valid": True, "skipped": True, "reason": reason}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["asset_res"]["duration"] = time.time() - step8_start
            self.step_result.emit(8, {"asset_res": v_res["asset_res"]})

            # 11.6 数据属性重复检测 (Node 9)
            step9_start = time.time()
            run_data_attr_check = raw_info.get("run_data_attribute_check", True)

            if "data_attr_res" in self._resumed_v_res:
                pass
            elif run_data_attr_check and pair["excel"] and os.path.exists(pair["excel"]):
                RuntimeLogger.log(f"正在进行 [Step 9] 数据属性重复检测...")
                self.progress.emit(98, "正在检测数据属性重复...")
                keywords = raw_info.get("data_attr_keywords", ["ERX", "EW", "EX"])
                split_sheet = raw_info.get("hierarchy_sheet")

                def data_attr_progress_proxy(current, total, msg):
                    if not self._is_running:
                        self._dump_pause_state()
                        return
                    if total > 0:
                        p = current / total * 100
                    else:
                        p = 0
                    mapped_progress = 98 + int(p * 0.01)
                    self.progress.emit(min(mapped_progress, 99), f"数据属性检测: {msg}")

                data_attr_res = check_data_attribute_duplicates(
                    file_path=pair["excel"], sheet_name=split_sheet,
                    keywords=keywords, progress_callback=data_attr_progress_proxy,
                    report_base_name=getattr(self, "_run_base_name", None)
                )
                if data_attr_res is None:
                    return
                v_res["data_attr_res"] = data_attr_res
                self._current_v_res = v_res  # [NEW] 立即保存
                RuntimeLogger.log(f"数据属性重复检测完成: {data_attr_res.get('statistics', {})}")
            else:
                if not run_data_attr_check:
                    reason = "未勾选数据属性重复检测节点"
                elif not pair["excel"]:
                    reason = "未检测到Excel文件"
                else:
                    reason = "Excel文件不存在"
                v_res["data_attr_res"] = {"success": True, "skipped": True, "reason": reason}

            if not self._is_running:
                self._dump_pause_state()
                return

            v_res["data_attr_res"]["duration"] = time.time() - step9_start
            self.step_result.emit(9, {"data_attr_res": v_res["data_attr_res"]})

            self.progress.emit(99, "正在进行最后的评估报告汇总...")

            # 12. 自动生成报表
            RuntimeLogger.log(f"正在汇总结果并生成评估报告...")
            report_path = ReportGenerator.generate_validation_report(
                self.task_data["filename"], v_res,
                base_name=getattr(self, "_run_base_name", None),
            )
            v_res["auto_report_path"] = report_path

            # 13. 保存运行日志
            log_path = RuntimeLogger.save_to_file(self.task_data["filename"])
            v_res["runtime_log_path"] = log_path

            self._clear_pause_state()
            self.progress.emit(100, "所有校验任务已完成")

            # 【新增】保存完整结果
            self._save_result_json(v_res)
            self._save_summary_html(v_res)

            self._finished_normally = True  # [NEW] 标记正常完成
            self.finished.emit(v_res)

        except Exception as e:
            import traceback
            RuntimeLogger.log(f"❌ 校验任务中断: {str(e)}", level="ERROR")
            RuntimeLogger.log(traceback.format_exc(), level="DEBUG")
            self.finished.emit({"is_valid": False, "error": str(e)})

        finally:
            # 无论成功失败，确保保存一次会话日志
            project_name = self.task_data.get("filename", "Unknown")
            try:
                RuntimeLogger.save_to_file(project_name)
                RuntimeLogger.close_file()
            except Exception as e:
                log_warn(f'[WARN] 保存日志失败: {e}')
            # [NEW] 兜底保存：如果任务非正常结束（崩溃/强制关闭），保存当前进度
            if not getattr(self, "_finished_normally", False):
                try:
                    log_info(f'[Worker] 任务非正常结束，尝试保存暂停状态...')
                    self._dump_pause_state()
                except Exception as e:
                    log_error(f'[ERROR] 兜底保存暂停状态失败: {e}')
                    import traceback
                    traceback.print_exc()

            # 任务完成后只清理本任务注册的临时文件（并行安全）
            try:
                DocumentProcessor.clear_cache(owner=RuntimeLogger.get_project())
            except:
                pass

            # 【新增】结束大任务标记并恢复线程优先级
            end_heavy_task(getattr(self, "_heavy_task", False), kind="initial")



class TaskCard(QFrame):
    """任务卡片，带实时计时和平滑进度条"""

    def __init__(self, task_data, parent=None):
        super().__init__(parent)
        self.task_data = task_data
        self.start_time = time.time()
        self.est_total = 45.0
        self.current_display_progress = 0
        self.target_backend_progress = 0
        self.current_step_num = 0
        self.is_running = True
        self.current_view_step = 0
        self._finished_cleanly = None  # None=运行中/手动停止, True=完成无异常, False=完成有异常
        self.setFrameShape(QFrame.StyledPanel)
        self.setProperty("class", "TaskCard")
        self._init_ui()
        self.update_style()

    def update_style(self):
        """全面刷新卡片样式"""
        from PySide6.QtWidgets import QApplication
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        left_bar_color = self.task_data.get("border_color", "#3b82f6")

        from utils.themes import get_tokens, rgba
        t = get_tokens(is_dark)
        card_bg = t["bg_card"]
        card_border = t["border"]
        text_color = t["text_sub"] if is_dark else t["text_body"]
        title_color = t["text_hi"]
        detail_card_bg = t["bg_card2"]
        val_text_color = t["text_hi"]
        lbl_text_color = t["text_sub"]
        btn_bg = t["bg_hover"] if is_dark else t["bg_card"]
        btn_hover = t["bg_hover"]
        btn_border = t["border"]
        btn_text = t["text_hi"]
        sep_color = t["border"]

        self.setStyleSheet(
            f"""
            QFrame[class="TaskCard"] {{
                background: {card_bg};
                border: 1px solid {card_border};
                border-left: 4px solid {left_bar_color};
                border-radius: 16px;
            }}
            QWidget#MainContainer {{ background: transparent; }}
            QWidget#CardInnerBox {{ background: transparent; }}
            QLabel {{ background: transparent; color: {text_color}; font-family: 'Segoe UI', 'Microsoft YaHei UI'; }}
            QLabel[class="task-title"] {{ color: {title_color}; font-size: 20px; font-weight: 700; }}
            QLabel[class="detail-title"] {{ color: {title_color}; font-size: 16px; font-weight: 800; }}
            QFrame#DetailCard {{ background: transparent; border: 1px solid {card_border}; border-radius: 12px; }}
            QTextBrowser {{ background: transparent; border: none; color: {text_color}; }}
            QPushButton {{
                padding: 6px 16px; border-radius: 8px; font-size: 13px; font-weight: 500;
                border: 1px solid {btn_border}; background-color: {btn_bg}; color: {btn_text};
            }}
            QPushButton:hover {{ background-color: {btn_hover}; border: 1px solid {t['accent']}; }}
            """
        )

        if hasattr(self, "stop_btn"):
            if self.is_running:
                stop_color = t["danger"]
            else:
                stop_color = t["success"] if self._finished_cleanly else t["danger"]
            self.stop_btn.setStyleSheet(f"color: {stop_color}; border-color: {rgba(stop_color, 0.4)};")

        if hasattr(self, "_meta_labels"):
            for lbl in self._meta_labels:
                lbl.setStyleSheet(f"color: {lbl_text_color}; font-size: 13px; font-weight: 500;")

        if hasattr(self, "_meta_values"):
            for val, orig_color in self._meta_values:
                target_color = val_text_color if orig_color == "DYNAMIC" else orig_color
                val.setStyleSheet(f"color: {target_color}; font-size: 18px; font-weight: 600;")

        if hasattr(self, "time_lbl"):
            self.time_lbl.setStyleSheet(f"color: {lbl_text_color}; font-size: 13px; font-weight: 500;")

        if hasattr(self, "time_label"):
            self.time_label.setStyleSheet(f"color: {val_text_color}; font-size: 13px; font-weight: 600;")

        if hasattr(self, "_separators"):
            for s in self._separators:
                s.setStyleSheet(f"color: {sep_color}; font-size: 13px; margin: 0 5px;")

        if hasattr(self, "badge"):
            s_color = t["success"]
            self.badge.setStyleSheet(
                f"background: {rgba(s_color, 0.1)}; color: {s_color}; "
                f"border-radius: 4px; padding: 2px 8px; font-weight: 700; "
                f"font-size: 11px; border: 1px solid {rgba(s_color, 0.2)};"
            )

        if hasattr(self, "steps_widget"):
            self.steps_widget.update_theme_style()

        if hasattr(self, "current_view_step") and self.current_view_step > 0:
            log_text = self.task_data["logs"].get(self.current_view_step, "")
            if log_text:
                self.update_log(self.current_view_step, log_text)

    def _create_metric_card(self, label, value, color="#3b82f6"):
        card = QFrame()
        card.setProperty("class", "MetricCard")
        card.setGraphicsEffect(None)
        layout = QVBoxLayout(card)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignCenter)
        lbl = QLabel(label)
        lbl.setProperty("class", "MetricLabel")
        lbl.setAlignment(Qt.AlignCenter)
        val = QLabel(str(value))
        val.setProperty("class", "MetricValue")
        val.setStyleSheet(f"color: {color};")
        val.setAlignment(Qt.AlignCenter)
        layout.addWidget(lbl)
        layout.addWidget(val)
        return card, val

    def _init_ui(self):
        from extend.matcher_config import MatcherConfig
        from utils.themes import get_tokens
        _tok = get_tokens(MatcherConfig.load().get("theme", {}).get("is_dark", False))

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        self.container_widget = QWidget()
        self.container_widget.setObjectName("MainContainer")
        layout = QVBoxLayout(self.container_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 16, 24, 16)
        main_layout.addWidget(self.container_widget)

        header = self._create_header()
        layout.addWidget(header)

        metrics_line = QHBoxLayout()
        metrics_line.setSpacing(15)

        def create_meta_item(label, value, color=None):
            container = QWidget()
            container.setObjectName("CardInnerBox")
            l = QHBoxLayout(container)
            l.setContentsMargins(0, 0, 0, 0)
            l.setSpacing(6)
            lbl = QLabel(f"{label}:")
            lbl.setProperty("class", "meta-label")
            val = QLabel(str(value))
            val.setProperty("class", "meta-value")
            l.addWidget(lbl)
            l.addWidget(val)
            if not hasattr(self, "_meta_labels"):
                self._meta_labels = []
            self._meta_labels.append(lbl)
            if not hasattr(self, "_meta_values"):
                self._meta_values = []
            self._meta_values.append((val, color if color else "DYNAMIC"))
            return container, val

        mandays_box, self.mandays_val = create_meta_item("送审人天", self.task_data.get("days", "0"))
        fp_box, self.fp_val = create_meta_item("送审工作量", "0", _tok["success"])

        self.time_label = QLabel("00:00")
        time_container = QWidget()
        time_container.setObjectName("CardInnerBox")
        time_layout = QHBoxLayout(time_container)
        time_layout.setContentsMargins(0, 0, 0, 0)
        time_layout.setSpacing(6)
        self.time_lbl = QLabel("总耗时:")
        self.time_lbl.setProperty("class", "meta-label")
        time_layout.addWidget(self.time_lbl)
        time_layout.addWidget(self.time_label)

        def create_sep():
            s = QLabel("|")
            s.setProperty("class", "metric-sep")
            if not hasattr(self, "_separators"):
                self._separators = []
            self._separators.append(s)
            return s

        metrics_line.addWidget(mandays_box)
        metrics_line.addWidget(create_sep())
        metrics_line.addWidget(fp_box)
        metrics_line.addWidget(create_sep())
        metrics_line.addWidget(time_container)
        metrics_line.addStretch()
        layout.addLayout(metrics_line)

        status_line = QHBoxLayout()
        status_line.setSpacing(8)
        self.status_icon = QLabel("🚀")
        self.status_icon.setStyleSheet("font-size: 14px;")
        self.status_label = QLabel("正在核查...")
        self.status_label.setStyleSheet(f"color: {_tok['accent']}; font-weight: 600; font-size: 14px;")
        status_line.addWidget(self.status_icon)
        status_line.addWidget(self.status_label)
        status_line.addStretch()
        layout.addLayout(status_line)

        self.steps_widget = StepsWidget(self.task_data["steps"])
        self.steps_widget.node_clicked.connect(self.on_node_clicked)
        self.steps_widget.label_clicked.connect(self.on_label_clicked)
        self.steps_widget.container.setFixedHeight(95)
        layout.addWidget(self.steps_widget)

        self.detail_card = QFrame()
        self.detail_card.setObjectName("DetailCard")
        self.detail_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(20, 15, 20, 15)
        detail_layout.setSpacing(10)

        detail_header = QHBoxLayout()
        self.step_badge = QLabel("STEP 01")
        self.step_badge.setStyleSheet(f"""
            background: {_tok['accent']}; color: white; border-radius: 4px;
            padding: 2px 8px; font-weight: 800; font-size: 10px;
        """)
        self.detail_title = QLabel("核查明细")
        self.detail_title.setProperty("class", "detail-title")

        detail_header.addWidget(self.step_badge)
        detail_header.addWidget(self.detail_title)
        detail_header.addStretch()
        detail_layout.addLayout(detail_header)

        self.detail_stack = QStackedWidget()
        self.detail_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.detail_content = QTextBrowser()
        self.detail_content.setOpenExternalLinks(True)
        from PySide6.QtGui import QTextOption
        self.detail_content.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        self.detail_content.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.detail_content.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.detail_content.setMinimumHeight(200)
        self.detail_content.setMaximumHeight(600)
        self.detail_stack.addWidget(self.detail_content)
        detail_layout.addWidget(self.detail_stack)
        layout.addWidget(self.detail_card)

        self.ui_timer = QTimer(self)
        self.ui_timer.timeout.connect(self.on_timer_tick)
        self.ui_timer.start(100)

        self.worker = None
        if not self.task_data.get("validation_results"):
            QTimer.singleShot(50, self.start_validation_worker)
        else:
            self.on_validation_finished(self.task_data["validation_results"][0])

    def start_validation_worker(self):
        """延迟启动校验 worker"""
        if self.worker is None:
            self.worker = ValidationWorker(self.task_data)
            self.worker.progress.connect(self.on_validation_progress)
            self.worker.step_result.connect(self.on_step_finished)
            self.worker.finished.connect(self.on_validation_finished)
            from utils.task_queue import TaskQueueManager
            started = TaskQueueManager().submit("initial", self.worker)
            if not started and hasattr(self, "status_label"):
                self.status_label.setText(f"⏳ {self._queue_hold_reason()}")
                self.status_icon.setText("🕘")

    @staticmethod
    def _queue_hold_reason():
        """排队原因细分：内存看门狗暂缓 / 大项目互斥 / 真在排队，不再一律显示
        "等待前面的任务完成"误导用户"""
        try:
            from utils.task_state import get_running_heavy_kinds
            heavy = get_running_heavy_kinds()
            if heavy:
                return f"大型项目({'、'.join(sorted(heavy))})执行中，本任务已自动排队"
            from utils.memory_guard import get_memory_status
            ok, desc = get_memory_status()
            if not ok:
                return f"内存占用较高，暂缓启动（每5秒自动重试）｜{desc}"
        except Exception:
            pass
        return "排队中，等待前面的任务完成..."

    def _create_header(self):
        from extend.matcher_config import MatcherConfig
        from utils.themes import get_tokens, rgba
        from utils.path_utils import clean_project_name
        _tok = get_tokens(MatcherConfig.load().get("theme", {}).get("is_dark", False))

        header = QWidget()
        header.setObjectName("CardInnerBox")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(15)

        display_name = clean_project_name(self.task_data["filename"])
        title_icon = QLabel("📄")
        title_icon.setStyleSheet("font-size: 18px;")
        title_text = ElidedTitleLabel(f"{display_name}")
        title_text.setProperty("class", "task-title")

        self.badge = QLabel("结算")
        self.badge.setStyleSheet(f"""
            background: {rgba(_tok['success'], 0.1)}; color: {_tok['success']};
            border-radius: 4px; padding: 2px 8px; font-weight: 700;
            font-size: 11px; border: 1px solid {rgba(_tok['success'], 0.2)};
        """)

        header_layout.addWidget(title_icon)
        header_layout.addWidget(title_text)
        header_layout.addWidget(self.badge)
        header_layout.addStretch()

        self.report_btn = QPushButton("结果汇总")
        self.report_btn.setCursor(Qt.PointingHandCursor)
        self.report_btn.clicked.connect(self.show_summary)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.clicked.connect(self.stop_task)
        header_layout.addWidget(self.report_btn)
        header_layout.addWidget(self.stop_btn)

        return header

    def stop_task(self):
        """停止当前校验任务"""
        if hasattr(self, "worker") and self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.is_running = False
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("已停止")
            self.time_label.setText("⏱️ 校验已手动停止")
            self.update_log(
                self.current_step_num,
                "🛑 用户手动停止了校验任务。\n"
                "⏸ 当前进度会在当前步骤结束后自动保存，"
                "可到「历史项目」页找到本条记录，点击『▶️ 继续执行』从下一步继续。",
            )
            self.steps_widget.set_step_status(self.current_step_num, "warn")
            self.ui_timer.stop()

    def on_timer_tick(self):
        """每秒平滑增长进度并更新计时器"""
        if not self.is_running:
            return
        elapsed = time.time() - self.start_time

        if self.current_display_progress > 5 and self.current_display_progress < 100:
            real_est_total = elapsed / (self.current_display_progress / 100.0)
            self.est_total = self.est_total * 0.95 + real_est_total * 0.05

        rem = max(1, self.est_total - elapsed)
        if self.current_display_progress < 95 and rem < 3:
            rem = 3

        self.time_label.setText(
            f"⏱️ 预计剩余 {int(rem // 60):02d}:{int(rem % 60):02d} | 已耗时 {int(elapsed // 60):02d}:{int(elapsed % 60):02d}"
        )

        if self.current_display_progress < self.target_backend_progress:
            diff = self.target_backend_progress - self.current_display_progress
            if diff > 15:
                self.current_display_progress += 0.6
            elif diff > 5:
                self.current_display_progress += 0.25
            else:
                self.current_display_progress += 0.08
        elif self.current_display_progress < 99.8:
            self.current_display_progress += 0.003

        self.steps_widget.set_total_progress(self.current_display_progress)

        ranges = [
            (0, 15), (15, 30), (30, 33), (33, 36), (36, 39),
            (39, 70), (70, 95), (95, 100),
        ]
        idx = max(0, min(len(ranges) - 1, self.current_step_num))
        curr_range = ranges[idx]
        span = curr_range[1] - curr_range[0]
        if span <= 0:
            span = 1
        relative = (self.current_display_progress - curr_range[0]) / span
        step_progress = int(max(0, min(100, relative * 100)))
        self.steps_widget.set_step_progress(self.current_step_num, step_progress)

    def _show_in_app_toast(self, title, msg):
        """[NEW] 应用内提醒"""
        win = self.window()
        toast = QLabel(win)
        toast.setText(f"  {title} — {msg}  ")
        toast.setWordWrap(True)
        toast.setMaximumWidth(380)
        toast.setStyleSheet(
            "background: rgba(17,24,39,235); color: #f8fafc;"
            "border: 1px solid #3b82f6; border-radius: 8px;"
            "padding: 10px 16px; font-size: 13px;"
        )
        toast.adjustSize()
        toast.move(max(10, win.width() - toast.width() - 30), 70)
        toast.show()
        toast.raise_()
        if not hasattr(self, "_toasts"):
            self._toasts = []
        self._toasts.append(toast)
        QTimer.singleShot(4000, toast.deleteLater)

    def _notify_finished(self, has_issue):
        """[NEW] 校验完成提醒 - 最终修复版"""
        try:
            config = MatcherConfig.load()
            auto_cfg = config.get("automation", {})
            flags = auto_cfg.get("notify_flags", {})

            # 兼容旧配置
            if not flags:
                mode = auto_cfg.get("notify_mode", "popup")
                flags = {"in_app": False, "popup": mode == "popup", "tray": mode == "tray"}

            # 如果三个都没选，直接返回
            if not any(flags.get(k) for k in ("in_app", "popup", "tray")):
                return

            win = self.window()
            title = "✅ 初评校验完成" if not has_issue else "❌ 初评校验完成（存在异常）"
            msg = f"《{self.task_data.get('filename', '')}》校验完成，请查看结果。"

            # ================= 1. 任务栏闪烁 (独立逻辑) =================
            # 只要任务完成了，且窗口不在最前面，就闪烁任务栏提醒用户
            # 这样即使你没勾选“应用内提醒”，任务栏也会闪，防止用户错过
            if win and not win.isActiveWindow():
                QApplication.alert(win, 2000)  # 闪烁 2 秒

            # ================= 2. 应用内提醒 (右上角小条) =================
            # 只有勾选了“应用内提醒”，才显示右上角黑色小条
            if flags.get("in_app"):
                try:
                    self._show_in_app_toast(title, msg)
                except Exception as e:
                    log_error(f'[Toast Error] {e}')

            # ================= 3. 桌面通知 (Windows 右下角) =================
            # 只有勾选了“桌面通知”，才显示系统通知
            if flags.get("tray"):
                try:
                    from PySide6.QtWidgets import QSystemTrayIcon
                    from PySide6.QtGui import QIcon

                    if getattr(self, "_tray", None) is None:
                        # 强制加载 logo.ico，解决显示 "C" 的问题
                        icon_path = get_resource_path("ui/logo.ico")
                        if not os.path.exists(icon_path):
                            icon_path = get_resource_path("ui/logo.png")

                        tray_icon = QIcon(icon_path) if os.path.exists(icon_path) else None
                        if tray_icon and not tray_icon.isNull():
                            self._tray = QSystemTrayIcon(tray_icon, win)
                        else:
                            self._tray = QSystemTrayIcon(win.windowIcon(), win)

                        self._tray.setToolTip("COSMIC 智能评估")
                        self._tray.show()  # 必须 show 才能发通知

                    # 发送通知
                    msg_type = QSystemTrayIcon.Warning if has_issue else QSystemTrayIcon.Information
                    self._tray.showMessage(title, msg, msg_type, 6000)
                except Exception as e:
                    log_error(f'[Tray Error] {e}')
                # 注意：这里去掉了 return，允许同时显示弹框

            # ================= 4. 弹框通知 (中间的大窗口) =================
            # 只有勾选了“弹框通知”，才显示截图里那个大弹框
            if flags.get("popup"):
                try:
                    from PySide6.QtWidgets import QDialog, QPushButton
                    from PySide6.QtGui import QIcon  # 【修复】此前 import 在 tray 分支内，
                    # tray 未勾选时弹窗分支会报 UnboundLocalError 导致完成弹窗不显示

                    # 关闭旧的弹框
                    if getattr(self, "_notify_dialog", None) is not None:
                        try:
                            self._notify_dialog.close()
                        except:
                            pass

                    # 创建新弹框
                    dlg = QDialog(win)
                    dlg.setWindowTitle("任务完成提醒")

                    # 【关键修复】给弹窗设置图标，解决标题显示 "C" 的问题
                    icon_path = get_resource_path("ui/logo.ico")
                    if os.path.exists(icon_path):
                        dlg.setWindowIcon(QIcon(icon_path))

                    dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowStaysOnTopHint)
                    dlg.setModal(False)
                    dlg.resize(420, 150)

                    lay = QVBoxLayout(dlg)
                    title_label = QLabel(title)
                    title_label.setStyleSheet("font-size: 15px; font-weight: 700; color: #2563eb;")
                    lay.addWidget(title_label)

                    name_label = QLabel(msg)
                    name_label.setWordWrap(True)
                    lay.addWidget(name_label)

                    btn_row = QHBoxLayout()
                    btn_row.addStretch()
                    btn_view = QPushButton("查看结果")
                    btn_view.setStyleSheet(
                        "background: #2563eb; color: white; border-radius: 4px; padding: 5px 15px; font-weight: bold;")
                    btn_close = QPushButton("关闭")
                    btn_row.addWidget(btn_view)
                    btn_row.addWidget(btn_close)
                    lay.addLayout(btn_row)

                    def _view():
                        dlg.close()
                        win.showNormal()
                        win.activateWindow()
                        self.show_summary()

                    btn_view.clicked.connect(_view)
                    btn_close.clicked.connect(dlg.close)
                    QTimer.singleShot(15000, dlg.close)

                    self._notify_dialog = dlg
                    dlg.show()
                except Exception as e:
                    log_error(f'[Popup Error] {e}')

        except Exception as e:
            log_error(f'[Notify Error] {e}')

    def on_validation_progress(self, value, sub_step_text=""):
        """后台报告真实进度点"""
        if not self.is_running:
            return
        old_step = self.current_step_num
        self.target_backend_progress = float(value)

        step_names = [
            "环境解析与预加载", "模板合规性校验", "Excel 空值扫描", "送审比例计算",
            "附加值因子提取", "核心层级关系校验", "简单过程内容匹配",
            "数据移动类型校验", "资产清单匹配", "数据属性重复检测",
        ]

        if value < 15:
            current_idx = 0
            self.current_step_num = 0
        elif value < 30:
            current_idx = 1
            self.current_step_num = 1
        elif value < 33:
            current_idx = 2
            self.current_step_num = 2
        elif value < 36:
            current_idx = 3
            self.current_step_num = 3
        elif value < 39:
            current_idx = 4
            self.current_step_num = 4
        elif value < 70:
            current_idx = 5
            self.current_step_num = 5
        elif value < 95:
            current_idx = 6
            self.current_step_num = 6
        elif value < 96:
            current_idx = 7
            self.current_step_num = 7
        elif value < 98:
            current_idx = 8
            self.current_step_num = 8
        elif value < 99:
            current_idx = 9
            self.current_step_num = 9
        else:
            current_idx = 9
            self.current_step_num = 9

        if hasattr(self, "status_label"):
            disp_step = self.current_step_num
            step_text = step_names[current_idx]
            display_text = sub_step_text if sub_step_text else step_text

            if "[PROCESS]" in display_text:
                clean_detail = display_text.replace("[PROCESS]", "").replace("过程匹配:", "").strip()
                if self.current_step_num == 6:
                    self.status_label.setText(f"正在进行第 {disp_step} 步:功能过程匹配 {clean_detail}")
                else:
                    self.status_label.setText(f"正在进行第 {disp_step} 步 - {step_text} {clean_detail}")
            else:
                self.status_label.setText(f"正在进行第 {disp_step} 步 - {display_text}...")

            self.status_icon.setText("")
            if value >= 100:
                self.status_label.setText("所有校验任务已完成")
                self.status_icon.setText("✅")

        if value < 100:
            current_log = self.task_data["logs"].get(self.current_step_num, "")
            if not (current_log.startswith("✅") or current_log.startswith("❌") or current_log.startswith(
                    "⚪") or "⚠️" in current_log):
                if sub_step_text:
                    log_msg = f"🔍 {sub_step_text}"
                    self.task_data["logs"][self.current_step_num] = log_msg
                    self.update_log(self.current_step_num, log_msg)
                else:
                    self.update_log(self.current_step_num, f"🔍 正在执行校验逻辑... ({int(value)}%)")

        if self.current_step_num > old_step:
            for s in range(0, self.current_step_num):
                curr_status = self.steps_widget.step_nodes[s].status
                if curr_status not in ["done", "fail", "warn", "finished", "skipped"]:
                    self.steps_widget.set_step_status(s, "finished")
            self.steps_widget.set_step_progress(self.current_step_num, 5)

    def on_step_finished(self, step_num, result_part):
        """当单个节点完成时，立即更新该节点的状态和日志"""
        if not self.task_data.get("validation_results"):
            self.task_data["validation_results"] = [{}]
        results = self.task_data["validation_results"][0]
        results.update(result_part)
        self._update_specific_step_ui(step_num, results)

    def _update_specific_step_ui(self, step_num, results):
        """更新特定步骤的 UI 状态和日志"""
        if step_num == 0:
            env_res = results.get("env_check", {})
            status = "done"
            if results.get("tree_text"):
                self.tree_text = results["tree_text"]
            log = f"✅ 环境准备完成，成功解析 {env_res.get('sections', 0)} 个章节内容（点击节点可查看结构树）。"
            self.steps_widget.set_step_status(0, status)
            self.task_data["logs"][0] = log
            self.update_log(0, log)

        elif step_num == 1:
            if results.get("skipped"):
                status = "skipped"
                log = "⚪ 模板校验已跳过。"
            else:
                suspect_count = len(results.get("suspect_modules", []))
                if results.get("is_valid") and suspect_count == 0:
                    status = "done"
                    log = "✅ 模板校验通过：全层级正文已填充。"
                elif suspect_count > 0:
                    status = "fail"
                    log = f"❌ 发现 {suspect_count} 处正文与模板高度相似，疑似未填写！"
                else:
                    status = "fail"
                    log = "❌ 模板校验发现缺失或严重不符。"
            dur = results.get("duration", 0)
            log += f" (耗时: {dur:.1f}s)"
            self.steps_widget.set_step_status(1, status)
            self.task_data["logs"][1] = log
            self.update_log(1, log)

        elif step_num == 2:
            excel_res = results.get("excel_check", {})
            if excel_res.get("skipped"):
                status = "skipped"
                log = "⚪ 空值检查已跳过。"
            else:
                status = "done"
                log = "✅ Excel 关键列空值检查通过。"
                if not excel_res.get("is_ok"):
                    errors = excel_res.get("errors", [])
                    if errors:
                        status = "fail"
                        log = "❌ Excel 发现空值行：\n- " + "\n- ".join(errors[:5])
                        if len(errors) > 5:
                            log += f"\n...等共 {len(errors)} 项异常"
            dur = excel_res.get("duration", 0)
            log += f" (耗时: {dur:.1f}s)"
            self.steps_widget.set_step_status(2, status)
            self.task_data["logs"][2] = log
            self.update_log(2, log)

        elif step_num == 3:
            ratio_res = results.get("ratio_check", {})
            if ratio_res:
                if ratio_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 送审比例校验已跳过。"
                else:
                    status = "done" if ratio_res.get("is_ok") else "fail"
                    mandays = ratio_res.get("mandays", 0)
                    upper_limit = "2.0" if mandays <= 1000 else "1.5"
                    if ratio_res.get("is_ok"):
                        log = f"✅送审比例正常，当前送审比例为【{ratio_res.get('ratio', 0)}】,送审功能点'：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"
                    else:
                        ratio_val = ratio_res.get("ratio", 0)
                        desc = "过多" if ratio_val >= float(upper_limit) else "过少"
                        log = f"❌送审比例{desc}，当前送审比例为【{ratio_val}】，送审功能点：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"
                    if ratio_res.get("is_ok"):
                        log += f"\n- 送审比例范围在 0.8 ~ {upper_limit}"
                dur = ratio_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                fp_count = ratio_res.get("fp_count")
                if fp_count is not None:
                    self.fp_val.setText(str(fp_count))
                self.steps_widget.set_step_status(3, status)
                self.task_data["logs"][3] = log
                self.update_log(3, log)

        elif step_num == 4:
            factors = results.get("factor_check", {})
            if factors:
                if factors.get("skipped"):
                    status = "skipped"
                    log = " 附加值因子校验已跳过。"
                else:
                    report_lines = []
                    scale = factors.get("scale", {})
                    scale_val = scale.get("value")
                    if scale_val == "结算":
                        report_lines.append(f"✅ 需求变更规模因子: {scale_val}")
                    else:
                        report_lines.append(
                            f"❌ 需求变更规模因子: {scale_val if scale_val else '无'} (一般为结算, 请确认)")

                    if scale_val in ["结算", "预算"]:
                        self.badge.setText(scale_val)
                        badge_style = {
                            "结算": "background: #ecfdf5; color: #047857; border: 1px solid #6ee7b7;",
                            "预算": "background: #fffbeb; color: #b45309; border: 1px solid #fcd34d;",
                        }
                        self.badge.setStyleSheet(
                            f"{badge_style.get(scale_val)} padding: 2px 8px; border-radius: 4px; font-size: 11px;")

                    quality_keys = ["distributed", "performance", "reliability", "multiple_sites"]
                    name_map = {
                        "distributed": "分布式处理", "performance": "性能",
                        "reliability": "可靠性", "multiple_sites": "多重站点",
                    }
                    text_missing = []
                    text_ok_names = []
                    table_vals = []
                    table_missing_names = []
                    has_consistency_issue = False

                    for key in quality_keys:
                        f = factors.get(key, {})
                        name = name_map.get(key, key)
                        if f.get("found_in_text"):
                            text_ok_names.append(name)
                        else:
                            text_missing.append(name)
                        t_val = str(f.get("table_value") or "").strip()
                        if f.get("found_in_table") and t_val not in ["缺失", "-1"]:
                            table_vals.append(f"{name}({t_val if t_val in ['0', '1'] else '有描述'})")
                        else:
                            table_missing_names.append(name)
                        if f.get("consistency_warn"):
                            has_consistency_issue = True

                    if not text_missing:
                        report_lines.append(f"✅ 质量及特性文字描述: {'、'.join(text_ok_names)}")
                    else:
                        report_lines.append(f"❌ 质量及特性文字描述异常: {'、'.join(text_missing)} 缺少")

                    if not table_missing_names:
                        report_lines.append(f"✅ 质量及特性表格描述: {'、'.join(table_vals)}")
                    else:
                        all_t_parts = table_vals + [f"{m}(缺失)" for m in table_missing_names]
                        report_lines.append(f"❌ 质量及特性表格描述异常: {'、'.join(all_t_parts)}")

                    if not has_consistency_issue and not text_missing and not table_missing_names:
                        report_lines.append(f"✅ 质量及特性描述一致性: 正常")
                    else:
                        report_lines.append(f"❌ 质量及特性一致性校验不匹配")

                    log = "\n".join(report_lines)
                    status = "fail" if any(line.startswith("") for line in report_lines) else "done"

                self.steps_widget.set_step_status(4, status)
                self.task_data["logs"][4] = log
                self.update_log(4, log)

        elif step_num == 5:
            hierarchy_res = results.get("hierarchy_res", {})
            if hierarchy_res:
                if hierarchy_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 层级匹配校验已跳过。"
                else:
                    stats = hierarchy_res.get("statistics", {})
                    miss_count = stats.get("缺失项", 0)
                    mismatch_count = stats.get("层级不匹配", 0)
                    match_rate_str = stats.get("匹配率", "0%")
                    try:
                        match_rate_val = float(match_rate_str.strip("%")) / 100.0
                    except:
                        match_rate_val = 0
                    total_excel = stats.get("Excel功能点总数", 0)

                    if total_excel == 0:
                        status = "fail"
                        log = "❌ 层级匹配未执行：Excel 中未找到有效的三级模块数据。"
                    elif miss_count == 0 and mismatch_count == 0:
                        status = "done"
                        log = f"✅层级匹配通过（通过率{match_rate_str}）：所有Excel模块均在大纲中找到。"
                    else:
                        status = "fail"
                        not_found_items = hierarchy_res.get("not_found_in_word", [])
                        mismatched_items = hierarchy_res.get("hierarchy_mismatched", [])
                        all_failed_items = not_found_items + mismatched_items
                        issue_details = []

                        def _group_by_prefix(desc_list):
                            groups = {}
                            order = 0
                            for raw in desc_list:
                                if not raw or raw == "-":
                                    continue
                                desc = " ".join(str(raw).split())
                                parts = desc.split("：", 1)
                                prefix = parts[0].strip()
                                body = parts[1].strip() if len(parts) > 1 else ""
                                if prefix not in groups:
                                    groups[prefix] = {"order": order, "bodies": []}
                                    order += 1
                                if body and body not in groups[prefix]["bodies"]:
                                    groups[prefix]["bodies"].append(body)
                                elif not body and body not in groups[prefix]["bodies"]:
                                    groups[prefix]["bodies"].append(body)
                            return groups

                        miss_descs = []
                        for item in all_failed_items:
                            d = item.get("缺失简略描述")
                            if not d or d == "-":
                                d_gen = item.get("简略描述")
                                if d_gen and "未体现" in str(d_gen):
                                    d = d_gen
                            if d and d != "-":
                                miss_descs.append(d)

                        miss_groups = _group_by_prefix(miss_descs)
                        if miss_groups:
                            lines = []
                            for prefix, info in sorted(miss_groups.items(), key=lambda kv: kv[1]["order"]):
                                bodies = info["bodies"]
                                body0 = bodies[0] if bodies else ""
                                count = len([b for b in bodies if b]) or 1
                                if body0:
                                    lines.append(f"{prefix}：{body0}（已合并{count}条）")
                                else:
                                    lines.append(f"{prefix}（已合并{count}条）")
                            issue_details.append(f"• [缺失项] (Excel在Word未体现): \n  - " + "\n  - ".join(lines[:3]))

                        mismatch_descs = []
                        for item in all_failed_items:
                            d = item.get("层级不匹配简略描述")
                            if not d or d == "-":
                                d_gen = item.get("简略描述", "")
                                if d_gen and "不匹配" in str(d_gen):
                                    d = d_gen
                            if d and d != "-":
                                mismatch_descs.append(d)

                        mismatch_groups = _group_by_prefix(mismatch_descs)
                        if mismatch_groups:
                            lines = []
                            for prefix, info in sorted(mismatch_groups.items(), key=lambda kv: kv[1]["order"]):
                                bodies = info["bodies"]
                                body0 = bodies[0] if bodies else ""
                                count = len([b for b in bodies if b]) or 1
                                if body0:
                                    lines.append(f"{prefix}：{body0}（已合并{count}条）")
                                else:
                                    lines.append(f"{prefix}（已合并{count}条）")
                            issue_details.append(f"• [层级不匹配] (对应关系错误): \n  - " + "\n  - ".join(lines[:7]))

                        total_issues = len(not_found_items) + len(mismatched_items)
                        log = f"❌层级匹配存在异常（通过率{match_rate_str}，共{total_issues}处问题）：\n" + "\n".join(
                            issue_details)
                        log += "\n\n📂 [提示]：点击上方圆圈图标可直接打开详细的 Excel 匹配报告。"

                dur = hierarchy_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                self.steps_widget.set_step_status(5, status)
                self.task_data["logs"][5] = log
                self.update_log(5, log)

        elif step_num == 6:
            process_res = results.get("process_res", {})
            if process_res:
                if process_res.get("skipped"):
                    status = "skipped"
                    log = " 功能过程校验已跳过。"
                else:
                    stats = process_res.get("statistics", {})
                    miss_count = stats.get("缺失项", 0)
                    match_rate = stats.get("匹配率", "0%")
                    if miss_count == 0:
                        status = "done"
                        log = f"✅功能过程校验通过（匹配率{match_rate}）：所有功能过程描述均在正文中找到。"
                    else:
                        status = "fail"
                        not_found = process_res.get("not_found_in_word", [])
                        names = [f"【{item.get('Excel功能点', '未知')}】" for item in not_found[:2]]
                        names_str = "、".join(names)
                        suffix = "等" if len(not_found) > 2 else ""
                        log = f"❌功能过程校验不通过（匹配率：{match_rate}）\n{names_str}{suffix}功能过程在需求规格书未体现"
                    log += "\n\n📂 [提示]：点击上方圆圈图标可直接打开详细的 Excel 功能过程匹配报告。"
                dur = process_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                self.steps_widget.set_step_status(6, status)
                self.task_data["logs"][6] = log
                self.update_log(6, log)

        elif step_num == 7:
            move_res = results.get("move_res", {})
            if move_res:
                if move_res.get("skipped"):
                    status = "skipped"
                    log = " 数据移动类型校验已跳过。"
                else:
                    stats = move_res.get("statistics", {})
                    failed_count = stats.get("不合规", 0)
                    total_count = stats.get("总数", 0)
                    passed_count = stats.get("合规", 0)
                    if failed_count == 0 and total_count > 0:
                        status = "done"
                        log = f"✅数据移动类型校验通过（合规率 100%）：监测到 {total_count} 个功能过程，全部符合 E 开头、W/X 结束的规则。"
                    elif total_count == 0:
                        status = "fail"
                        log = "❌ 未发现有效的功能过程数据移动类型数据。"
                    else:
                        status = "fail"
                        failed_list = [r for r in move_res.get("items", []) if r.get("result") != "合规"]
                        err_types = {}
                        for item in failed_list:
                            err = item.get("result", "不合规")
                            if err not in err_types:
                                err_types[err] = []
                            err_types[err].append(item.get("process", "未知"))
                        err_details = []
                        for err, procs in err_types.items():
                            sample = "、".join(procs[:3]) + ("等" if len(procs) > 3 else "")
                            err_details.append(f"  - [{err}]: {sample} ({len(procs)}处)")
                        log = f"❌数据移动类型异常：共监测 {total_count} 处，其中 {failed_count} 处不合规。\n" + "\n".join(
                            err_details)
                        log += "\n\n📂 [提示]：请点击上方圆圈图标查看详细的 Excel 数据移动校验报告。"
                dur = move_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                self.steps_widget.set_step_status(7, status)
                self.task_data["logs"][7] = log
                self.update_log(7, log)

        elif step_num == 8:
            asset_res = results.get("asset_res", {})
            if asset_res:
                if asset_res.get("skipped"):
                    status = "skipped"
                    reason = asset_res.get("reason", "")
                    log = f"⚪ 资产清单匹配已跳过。{('原因: ' + reason) if reason else ''}"
                elif not asset_res.get("is_valid"):
                    status = "fail"
                    log = f"❌资产清单匹配执行出错：{asset_res.get('error', '未知错误')}"
                else:
                    stats = asset_res.get("statistics", {})
                    total_cnt = stats.get("总数", 0)
                    exact_cnt = stats.get("完全匹配", 0)
                    fuzzy_cnt = stats.get("模糊匹配", 0)
                    mismatch_cnt = stats.get("不匹配", 0)
                    match_rate_str = stats.get("匹配率", "0%")
                    if total_cnt == 0:
                        status = "fail"
                        log = "❌ 资产清单匹配未执行：拆分表中未找到有效的层级数据。"
                    elif mismatch_cnt == 0:
                        status = "done"
                        log = f"✅资产清单匹配通过（匹配率{match_rate_str}）：共 {total_cnt} 项，完全匹配 {exact_cnt}、模糊匹配 {fuzzy_cnt}，全部在资产清单中找到对应。"
                    else:
                        status = "fail"
                        items = asset_res.get("items", [])
                        missing_lines = []
                        mismatch_lines = []
                        for item in items:
                            if item.get("匹配状态") != "不匹配":
                                continue
                            desc = item.get("缺失简略描述", "")
                            if desc and desc != "无缺失" and desc not in missing_lines:
                                missing_lines.append(desc)
                            desc2 = item.get("层级不匹配简略描述", "")
                            if desc2 and desc2 != "无层级错位" and desc2 not in mismatch_lines:
                                mismatch_lines.append(desc2)
                        issue_details = []
                        if missing_lines:
                            issue_details.append(
                                "• [缺失项] (拆分表模块在资产清单无对应):\n  - " + "\n  - ".join(missing_lines[:5]))
                        if mismatch_lines:
                            issue_details.append("• [层级不匹配] (拆分表与资产清单层级不一致):\n  - " + "\n  - ".join(
                                mismatch_lines[:5]))
                        remaining = mismatch_cnt - (len(missing_lines) + len(mismatch_lines))
                        log = f"❌资产清单匹配存在异常（匹配率{match_rate_str}，共{mismatch_cnt}处问题）：\n" + "\n".join(
                            issue_details)
                        if remaining > 0:
                            log += f"\n  ...等共 {mismatch_cnt} 处问题"
                        log += "\n\n📂 [提示]：点击上方圆圈图标可直接打开详细的资产清单匹配 Excel 报告。"
                dur = asset_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                self.steps_widget.set_step_status(8, status)
                self.task_data["logs"][8] = log
                self.update_log(8, log)

        elif step_num == 9:
            data_attr_res = results.get("data_attr_res", {})
            if data_attr_res:
                if data_attr_res.get("skipped"):
                    status = "skipped"
                    reason = data_attr_res.get("reason", "")
                    log = f"⚪ 数据属性重复检测已跳过。{('原因: ' + reason) if reason else ''}"
                elif not data_attr_res.get("success"):
                    status = "fail"
                    log = f"❌数据属性重复检测执行出错：{data_attr_res.get('error', '未知错误')}"
                else:
                    stats = data_attr_res.get("statistics", {})
                    total_marked = stats.get("total_marked_rows", 0)
                    intra_dups = stats.get("intra_duplicates", 0)
                    cross_dups = stats.get("cross_duplicates", 0)
                    patterns_matched = stats.get("patterns_matched", 0)
                    if total_marked == 0:
                        status = "done"
                        log = f"✅数据属性重复检测通过：未发现任何重复项"
                    else:
                        status = "warn"
                        log = f"️数据属性重复检测发现 {total_marked} 行存在重复：\n"
                        log += f"  • 单功能过程内重复: {intra_dups} 组\n"
                        log += f"  • 跨功能过程重复: {cross_dups} 组（匹配 {patterns_matched} 个模式）\n"
                        log += f"📂 [提示]：点击上方圆圈图标可打开详细的重复检测报告 Excel"
                dur = data_attr_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"
                self.steps_widget.set_step_status(9, status)
                self.task_data["logs"][9] = log
                self.update_log(9, log)

        from utils.themes import get_tokens as _gt
        _tok = _gt(MatcherConfig.load().get("theme", {}).get("is_dark", False))
        has_any_fail = False
        for node in getattr(self.steps_widget, "step_nodes", []):
            if node.status in ["fail", "error", "warn"]:
                has_any_fail = True
                break

        if has_any_fail:
            if hasattr(self, "status_label"):
                self.status_label.setStyleSheet(f"font-weight: 800; font-size: 14px; color: {_tok['danger']};")
        else:
            if hasattr(self, "status_label"):
                self.status_label.setStyleSheet(f"font-weight: 600; font-size: 14px; color: {_tok['accent']};")

    def on_validation_finished(self, results):
        """校验完全结束"""
        self.is_running = False
        self.ui_timer.stop()
        from utils.themes import get_tokens as _gt, rgba as _rgba
        _tok = _gt(MatcherConfig.load().get("theme", {}).get("is_dark", False))
        if hasattr(self, "stop_btn"):
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("已完成")

        self.current_display_progress = 100
        if hasattr(self, "steps_widget"):
            self.steps_widget.set_total_progress(100)
            self.steps_widget.clear_all_progress()
            for i in range(1, 10):
                if i > len(self.steps_widget.step_nodes):
                    break
                node = self.steps_widget.step_nodes[i - 1]
                if node.status in ["pending", "processing"]:
                    self.steps_widget.set_step_status(i, "finished")

        has_issue = False
        if hasattr(self, "steps_widget"):
            for node in self.steps_widget.step_nodes:
                if node.status in ["fail", "error", "warn"]:
                    has_issue = True
                    break

        self._finished_cleanly = not has_issue
        if hasattr(self, "stop_btn"):
            # 完成态按钮不再是红色警示色：无异常为绿色，有异常保持红色
            _c = _tok["success"] if not has_issue else _tok["danger"]
            self.stop_btn.setStyleSheet(f"color: {_c}; border-color: {_rgba(_c, 0.4)};")

        status_text = "所有任务校验完成"
        if has_issue:
            status_text += " (存在异常)"
        self.status_label.setText(status_text)
        self.status_icon.setText("❌" if has_issue else "✅")

        from utils.themes import get_tokens as _gt
        _tok = _gt(MatcherConfig.load().get("theme", {}).get("is_dark", False))
        self.status_label.setStyleSheet(
            f"font-weight: 800; font-size: 14px; color: {_tok['danger'] if has_issue else _tok['success']};"
        )

        elapsed = time.time() - self.start_time
        self.time_label.setText(f"{int(elapsed // 60):02d}:{int(elapsed % 60):02d}")

        self._notify_finished(has_issue)

        if not results:
            return

        if not self.task_data.get("validation_results"):
            self.task_data["validation_results"] = [results]
        else:
            self.task_data["validation_results"][0].update(results)

        final_show_step = 7
        res_dict = self.task_data["validation_results"][0]
        if res_dict.get("excel_check", {}).get("is_ok") == False:
            final_show_step = 2
        elif res_dict.get("ratio_check", {}).get("is_ok") == False:
            final_show_step = 3
        elif res_dict.get("factor_check") and self.steps_widget.step_nodes[3].status != "done":
            final_show_step = 4
        elif res_dict.get("hierarchy_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 5
        elif res_dict.get("process_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 6
        elif res_dict.get("asset_res", {}).get("statistics", {}).get("不匹配", 0) > 0:
            final_show_step = 8
        elif res_dict.get("data_attr_res", {}).get("statistics", {}).get("total_marked_rows", 0) > 0:
            final_show_step = 9

        for i in range(1, 10):
            if i not in self.task_data["logs"]:
                self.task_data["logs"][i] = "✅ 校验通过，未发现异常。"

        final_log = self.task_data["logs"].get(final_show_step, "")
        self.update_log(final_show_step, final_log)

        if res_dict.get("excel_check", {}).get("is_ok") == False:
            final_show_step = 2
        elif res_dict.get("ratio_check", {}).get("is_ok") == False:
            final_show_step = 3
        elif res_dict.get("factor_check") and self.steps_widget.step_nodes[3].status != "done":
            final_show_step = 4
        elif (res_dict.get("hierarchy_res", {}).get("statistics", {}).get("缺失项", 0) > 0 or
              res_dict.get("hierarchy_res", {}).get("statistics", {}).get("层级不匹配", 0) > 0):
            final_show_step = 5
        elif res_dict.get("process_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 6
        elif res_dict.get("asset_res", {}).get("statistics", {}).get("不匹配", 0) > 0:
            final_show_step = 8
        elif res_dict.get("data_attr_res", {}).get("statistics", {}).get("total_marked_rows", 0) > 0:
            final_show_step = 9

        if final_show_step not in self.task_data["logs"]:
            for i in range(9, 0, -1):
                if i in self.task_data["logs"]:
                    final_show_step = i
                    break

        final_log = self.task_data["logs"].get(final_show_step, "")
        if res_dict.get("auto_report_path"):
            final_log += f"\n\n📂 完整评估报告已自动保存至：{res_dict.get('auto_report_path')}"
        self.update_log(final_show_step, final_log)

        config = MatcherConfig.load()
        if config.get("automation", {}).get("auto_open", True):
            report_path = res_dict.get("auto_report_path")
            if report_path:
                open_directory(report_path)

    def on_view_toggle(self, btn_id):
        """切换文字/表格视图 (已暂时注释)"""
        pass

    def _update_table_data(self, step_num):
        """根据当前步骤结果渲染高仿原型图的表格布局 (已暂时注释)"""
        pass

    def update_log(self, step_num, text):
        self.current_view_step = step_num
        from extend.matcher_config import MatcherConfig
        from utils.themes import get_tokens as _gt
        _tok = _gt(MatcherConfig.load().get("theme", {}).get("is_dark", False))
        color_map = {
            "processing": _tok["accent"],
            "success": _tok["success"],
            "error": _tok["danger"],
        }

        if "❌" in text or "⚠️" in text:
            current_color = color_map["error"]
        elif "✅" in text:
            current_color = color_map["success"]
        else:
            current_color = color_map["processing"]

        from extend.matcher_config import MatcherConfig
        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)

        border_col = "#334155" if is_dark else "#e2e8f0"
        self.detail_card.setStyleSheet(
            f"""
            QFrame#DetailCard {{
                background: transparent;
                border: 1px solid {border_col};
                border-left: 4px solid {current_color};
                border-radius: 10px;
            }}
            """
        )

        self.step_badge.setText(f"STEP {step_num:02d}")
        self.step_badge.setStyleSheet(
            f"""
            background: {current_color}; color: white; border-radius: 4px;
            padding: 2px 8px; font-weight: 800; font-size: 10px;
            """
        )

        step_name = "详情"
        if 0 <= step_num < len(self.task_data["steps"]):
            step_name = self.task_data["steps"][step_num][0]
        self.detail_title.setText(f"{step_name}校验报告")
        self.detail_title.setStyleSheet(f"font-weight: 800; font-size: 16px; color: {current_color};")

        formatted_text = text.replace("\n", "<br/>")
        text_col = "#cbd5e1" if is_dark else "#334155"
        style = f"color: {text_col}; background: transparent; font-family: 'Segoe UI', 'Microsoft YaHei UI'; font-size: 14px; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word;"
        self.detail_content.setHtml(f"<div style='{style}'>{formatted_text}</div>")

        if self.detail_stack.currentIndex() == 1:
            self._update_table_data(step_num)

    def on_label_clicked(self, step_num):
        """点击文字：仅展示日志"""
        log_text = self.task_data["logs"].get(step_num, "⏳ 等待处理...")
        self.update_log(step_num, log_text)

    def on_node_clicked(self, step_num):
        """点击圆圈：完成看报告，未完成看日志"""
        if step_num == 0:
            tree_txt_path = getattr(self.worker, "_tree_txt_path", None) if hasattr(self, "worker") else None
            if tree_txt_path and os.path.exists(tree_txt_path):
                try:
                    os.startfile(tree_txt_path)
                except Exception as e:
                    self.update_log(0, f"❌ 无法打开结构树文件: {e}")
                return
            tree_text = getattr(self, "tree_text", "")
            if tree_text:
                import tempfile
                tmp_txt = os.path.join(tempfile.gettempdir(), "Word结构树.txt")
                try:
                    with open(tmp_txt, "w", encoding="utf-8") as f:
                        f.write(tree_text)
                    os.startfile(tmp_txt)
                except Exception as e:
                    self.update_log(0, f"❌ 无法打开结构树文件: {e}")
                return
            self.on_label_clicked(step_num)

        elif step_num == 1:
            if not self.is_running and self.task_data.get("validation_results"):
                self.show_report()
            else:
                self.on_label_clicked(step_num)

        elif step_num == 2:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("excel_check", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(2, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        elif step_num == 5:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("hierarchy_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(5, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        elif step_num == 6:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("process_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(6, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        elif step_num == 7:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("move_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(7, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        elif step_num == 8:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("asset_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(8, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        elif step_num == 9:
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("data_attr_res", {}).get("output_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(9, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)

        else:
            self.on_label_clicked(step_num)

    def on_step_clicked(self, step_num):
        self.on_node_clicked(step_num)

    def _open_review_dir(self):
        """[NEW] 打开本项目的初评报告文件夹（直达最新一次运行的时间戳子目录）"""
        try:
            out_dir = latest_run_dir(self.task_data.get("filename", ""))
            os.makedirs(out_dir, exist_ok=True)
            open_directory(out_dir)
        except Exception as e:
            self.update_log(0, f"❌ 无法打开初评文件夹: {e}")

    def _open_log_dir(self):
        """[NEW] 打开本项目的日志文件夹"""
        try:
            log_dir = RuntimeLogger.get_log_dir(
                project_name=self.task_data.get("filename")
            )
            if not log_dir or not os.path.exists(log_dir):
                self.update_log(0, "⏳ 日志文件夹尚未生成（先运行一次校验）。")
                return
            open_directory(log_dir)
        except Exception as e:
            self.update_log(0, f"❌ 无法打开日志文件夹: {e}")

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        try:
            is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)
        except Exception:
            is_dark = False

        if is_dark:
            menu.setStyleSheet(
                """
                QMenu { background-color: #1e293b; color: #e2e8f0; border: 1px solid #334155; border-radius: 6px; padding: 4px; }
                QMenu::item { padding: 6px 24px 6px 12px; border-radius: 4px; }
                QMenu::item:selected { background-color: #3b82f6; color: #ffffff; }
                QMenu::separator { height: 1px; background: #334155; margin: 4px 8px; }
                """
            )

        view_report_action = QAction(" 查看详细报告", self)
        view_report_action.triggered.connect(self.show_report)
        menu.addAction(view_report_action)

        open_log_action = QAction("📂 打开运行日志", self)
        open_log_action.triggered.connect(self.open_runtime_log)
        menu.addAction(open_log_action)

        open_review_dir_action = QAction(" 打开初评文件夹", self)
        open_review_dir_action.triggered.connect(self._open_review_dir)
        menu.addAction(open_review_dir_action)

        open_log_dir_action = QAction("📜 打开日志文件夹", self)
        open_log_dir_action.triggered.connect(self._open_log_dir)
        menu.addAction(open_log_dir_action)

        copy_name_action = QAction("📋 复制项目名称", self)
        copy_name_action.triggered.connect(lambda: QApplication.clipboard().setText(self.task_data["filename"]))
        menu.addAction(copy_name_action)

        menu.exec(event.globalPos())

    def open_runtime_log(self):
        results = self.task_data.get("validation_results", [{}])[0]
        log_path = results.get("runtime_log_path")
        if log_path and os.path.exists(log_path):
            try:
                os.startfile(log_path)
            except Exception as e:
                self.update_log(1, f"❌ 无法打开日志文件: {e}")
        else:
            self.update_log(1, "⏳ 日志文件尚不存在或任务尚未完成。")

    def show_report(self):
        if not self.task_data.get("validation_results"):
            return
        dialog = ReportDialog(self.task_data, self)
        dialog.exec()

    def show_summary(self):
        dialog = SummaryDialog(self.task_data, self)
        dialog.exec()