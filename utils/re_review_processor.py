import os
import re
import shutil
import openpyxl
from copy import copy
from PySide6.QtCore import QThread, Signal
from utils.path_utils import get_resource_path
from utils.runtime_logger import log_info
from utils.task_state import begin_heavy_task_if_large, end_heavy_task


def _get_unique_copy_path(src_path, target_dir):
    """获取不重名的复制目标路径，如果重名自动添加 _1, _2 序号"""
    base_name = os.path.basename(src_path)
    name, ext = os.path.splitext(base_name)
    target_path = os.path.join(target_dir, base_name)

    if not os.path.exists(target_path):
        return target_path

    counter = 1
    while True:
        new_name = f"{name}_{counter}{ext}"
        target_path = os.path.join(target_dir, new_name)
        if not os.path.exists(target_path):
            return target_path
        counter += 1


class ReReviewWorker(QThread):
    """重评工作线程：三阶段处理"""
    progress = Signal(int)
    finished = Signal(tuple)
    error = Signal(str)

    def __init__(self, excel1_path, excel2_path, output_dir, project_name=None, target_sheet_name=None):
        super().__init__()
        self.excel1_path = excel1_path
        self.excel2_path = excel2_path
        self.output_dir = output_dir
        self.project_name = project_name or "ReReview"
        self.target_sheet_name = target_sheet_name

    def run(self):
        from utils.runtime_logger import RuntimeLogger
        from utils.evaluation_processor import EvaluationProcessor

        RuntimeLogger.set_project(self.project_name)
        RuntimeLogger.log("🚀 开始重评处理...")

        # 【新增】大项目检测：输入报告超阈值则提升线程优先级并暂缓历史页扫描
        self._heavy_task = begin_heavy_task_if_large(
            [self.excel1_path, self.excel2_path], self.project_name, kind="re_review"
        )

        try:
            # ================= 阶段 1：智能搬运数据生成新评估报告 =================
            RuntimeLogger.log("📂 [阶段1/3] 正在将新拆分表搬运到通用模板...")

            template_path = None
            possible_paths = [
                get_resource_path("folder/xxxx-评估报告.xlsx"),
                os.path.join(os.path.dirname(__file__), "../folder/xxxx-评估报告.xlsx"),
            ]
            for p in possible_paths:
                if p and os.path.exists(p):
                    template_path = p
                    break

            if not template_path:
                raise FileNotFoundError("找不到通用评估报告模板！")

            wb = openpyxl.load_workbook(template_path)
            ws_split = EvaluationProcessor.find_sheet_by_names(wb, ["功能点拆分表", "2、功能点拆分表", "拆分表"])
            if not ws_split:
                raise ValueError("通用模板中未找到'功能点拆分表'工作表")

            # 优先使用 UI 传来的 sheet，如果没有则后台智能兜底
            best_sheet_name = self.target_sheet_name
            if not best_sheet_name:
                best_sheet_name = ReReviewProcessor._get_best_sheet_name(self.excel2_path)
                if best_sheet_name:
                    RuntimeLogger.log(f"🎯 UI未指定，后台智能识别并锁定拆分表工作表: '{best_sheet_name}'")
            else:
                RuntimeLogger.log(f"🎯 使用用户指定的工作表: '{best_sheet_name}'")

            # 调用搬运方法（需确保 evaluation_processor.py 中的 _migrate_split_table_data 支持 target_sheet_name 参数）
            success = EvaluationProcessor._migrate_split_table_data(
                self.excel2_path,
                ws_split,
                lambda msg: RuntimeLogger.log(msg),
                # target_sheet_name=best_sheet_name
            )
            if not success:
                raise RuntimeError("数据搬运失败，请检查新拆分表格式")

            if self.progress:
                self.progress.emit(20)

            os.makedirs(self.output_dir, exist_ok=True)
            split_name = os.path.splitext(os.path.basename(self.excel2_path))[0]
            if split_name.endswith("-评估报告") or split_name.endswith("_评估报告"):
                split_name = re.sub(r"[-_]评估报告$", "", split_name)

            new_report_name = f"{split_name}-评估报告.xlsx"
            new_report_path = os.path.join(self.output_dir, new_report_name)
            wb.save(new_report_path)
            wb.close()

            RuntimeLogger.log(f"✅ 新报告已生成: {new_report_name}")
            if self.progress:
                self.progress.emit(35)

            # ================= 阶段 2：VLOOKUP 对比 + 模块颜色同步 =================
            RuntimeLogger.log("🔄 [阶段2/3] 正在对比新旧报告，同步模块颜色并生成差异标注...")

            old_annotated, new_annotated = ReReviewProcessor.process_re_review(
                self.excel1_path,
                new_report_path,
                self.output_dir,
                self.project_name,
                lambda v: self.progress.emit(35 + int(v * 0.50)) if self.progress else None
            )

            RuntimeLogger.log("✅ 差异标注完成！")
            if self.progress:
                self.progress.emit(90)

            # ================= 阶段 3：智能复制到上传拆分表所在目录 =================
            excel2_dir = os.path.dirname(os.path.abspath(self.excel2_path))
            output_dir_abs = os.path.abspath(self.output_dir)

            if excel2_dir != output_dir_abs:
                copy_path = _get_unique_copy_path(new_annotated, excel2_dir)
                shutil.copy2(new_annotated, copy_path)
                RuntimeLogger.log(f"📂 已将新报告副本保存至上传目录: {os.path.basename(copy_path)}")
            else:
                RuntimeLogger.log("ℹ️ 上传目录与输出目录相同，跳过复制。")

            if self.progress:
                self.progress.emit(100)

            RuntimeLogger.log("🎉 重评任务全部完成！")
            self.finished.emit((old_annotated, new_annotated))

        except Exception as exc:
            import traceback
            traceback.print_exc()
            self.error.emit(str(exc))
        finally:
            # 【新增】结束大任务标记并恢复线程优先级
            end_heavy_task(getattr(self, "_heavy_task", False), kind="re_review")


class ReReviewProcessor:
    """Create re-review annotations between the old report and new receipt."""

    START_ROW = 5
    CURRENT_REUSE_COL = 20  # T
    NEW_ID_COL = 21  # U
    PREVIOUS_REUSE_COL = 16  # P
    PREVIOUS_RESULT_COL = 17  # Q
    REMARK_COL = 15  # O

    @staticmethod
    def _get_best_sheet_name(file_path):
        """智能选择拆分表工作表"""
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            if not sheet_names: return None
            for name in sheet_names:
                if "2、" in name and any(kw in name for kw in ["拆分", "功能点"]): return name
            for name in sheet_names:
                if any(kw in name for kw in ["功能点拆分表", "2、功能点拆分表", "拆分表", "功能点"]): return name
            return sheet_names[0]
        except Exception:
            return None

    @staticmethod
    def find_column(ws, names, default_col, search_rows=10):
        if isinstance(names, str): names = [names]
        for row in range(1, search_rows + 1):
            for col in range(1, ws.max_column + 1):
                value = ws.cell(row=row, column=col).value
                if value is not None and any(name in str(value) for name in names):
                    return col
        return default_col

    @staticmethod
    def _split_sheet(wb):
        exact_names = ["功能点拆分表", "2、功能点拆分表", "拆分表"]
        for name in exact_names:
            if name in wb.sheetnames: return wb[name]
        for name in wb.sheetnames:
            if "功能点拆分" in name or "拆分表" in name: return wb[name]
        raise ValueError("无法在文件中找到功能点拆分表工作表")

    @classmethod
    def _sync_module_styles(cls, ws_old, ws_new, l1_col_old, l2_col_old, l3_col_old, l1_col_new, l2_col_new,
                            l3_col_new):
        """将旧报告中一二三级模块的背景颜色同步到新报告中"""
        old_styles = {}
        for level, col in [(1, l1_col_old), (2, l2_col_old), (3, l3_col_old)]:
            current_name = ""
            for row in range(cls.START_ROW, ws_old.max_row + 1):
                val = ws_old.cell(row=row, column=col).value
                if val and str(val).strip():
                    current_name = str(val).strip()
                    fill = ws_old.cell(row=row, column=col).fill
                    if fill and fill.start_color and fill.start_color.rgb and fill.start_color.rgb != "00000000":
                        old_styles[(level, current_name)] = copy(fill)
        synced_count = 0
        for level, col in [(1, l1_col_new), (2, l2_col_new), (3, l3_col_new)]:
            current_name = ""
            for row in range(cls.START_ROW, ws_new.max_row + 1):
                val = ws_new.cell(row=row, column=col).value
                if val and str(val).strip():
                    current_name = str(val).strip()
                    key = (level, current_name)
                    if key in old_styles:
                        ws_new.cell(row=row, column=col).fill = old_styles[key]
                        synced_count += 1
        return synced_count

    @classmethod
    def _build_rows(cls, ws, func_col, desc_col):
        rows = {}
        current_process = ""
        occurrences = {}
        for row in range(cls.START_ROW, ws.max_row + 1):
            process_value = ws.cell(row=row, column=func_col).value
            if process_value is not None and str(process_value).strip():
                current_process = str(process_value).strip()
            description_value = ws.cell(row=row, column=desc_col).value
            if description_value is None or not str(description_value).strip():
                continue
            description = str(description_value).strip()
            base_key = (current_process, description)
            occurrences[base_key] = occurrences.get(base_key, 0) + 1
            rows[row] = {
                "key": (
                    f"{len(current_process)}:{current_process}{len(description)}:{description}#{occurrences[base_key]}"),
                "process": current_process,
                "description": description,
            }
        return rows

    @staticmethod
    def _replace_first_number(value, replacement):
        text = str(value).strip()
        match = re.search(r"\d+", text)
        if not match: return text
        start, end = match.span()
        if text == match.group(): return replacement
        return text[:start] + str(replacement) + text[end:]

    @classmethod
    def process_re_review(cls, excel1_path, excel2_path, output_dir, project_name, progress_callback=None):
        if progress_callback: progress_callback(5)

        wb1_formula = openpyxl.load_workbook(excel1_path)
        wb1_value = openpyxl.load_workbook(excel1_path, data_only=True)
        wb2_value = openpyxl.load_workbook(excel2_path, data_only=True)
        wb2_final = openpyxl.load_workbook(excel2_path)
        try:
            if progress_callback: progress_callback(25)

            ws1_formula = cls._split_sheet(wb1_formula)
            ws1_value = cls._split_sheet(wb1_value)
            ws2_value = cls._split_sheet(wb2_value)
            ws2_final = cls._split_sheet(wb2_final)

            # ================= 同步一二三级模块颜色 =================
            l1_col_old = cls.find_column(ws1_value, ["一级模块", "一级", "L1"], 2)
            l2_col_old = cls.find_column(ws1_value, ["二级模块", "二级", "L2"], 3)
            l3_col_old = cls.find_column(ws1_value, ["三级模块", "三级", "L3"], 4)
            l1_col_new = cls.find_column(ws2_final, ["一级模块", "一级", "L1"], 2)
            l2_col_new = cls.find_column(ws2_final, ["二级模块", "二级", "L2"], 3)
            l3_col_new = cls.find_column(ws2_final, ["三级模块", "三级", "L3"], 4)

            synced_count = cls._sync_module_styles(ws1_value, ws2_final, l1_col_old, l2_col_old, l3_col_old, l1_col_new,
                                                   l2_col_new, l3_col_new)
            if synced_count > 0: log_info(f'✅ 成功同步 {synced_count} 个模块的颜色标注到新报告')

            # ================= 原有 VLOOKUP 对比逻辑 =================
            desc_col1 = cls.find_column(ws1_value, ["子过程描述", "功能点拆分描述", "子过程"], 8)
            func_col1 = cls.find_column(ws1_value, ["功能过程", "功能简述", "功能点"], 6)
            desc_col2 = cls.find_column(ws2_value, ["子过程描述", "功能点拆分描述", "子过程"], 8)
            func_col2 = cls.find_column(ws2_value, ["功能过程", "功能简述", "功能点"], 6)

            excel1_rows = cls._build_rows(ws1_value, func_col1, desc_col1)
            excel2_rows = cls._build_rows(ws2_value, func_col2, desc_col2)

            def group_key(data):
                return data["process"], data["description"]

            old_groups = {}
            new_groups = {}
            for row, data in excel1_rows.items(): old_groups.setdefault(group_key(data), []).append(row)
            for row, data in excel2_rows.items(): new_groups.setdefault(group_key(data), []).append(row)

            new_to_old = {}
            assigned_old_rows = set()
            for new_row, new_data in excel2_rows.items():
                candidates = old_groups.get(group_key(new_data), [])
                available = [row for row in candidates if row not in assigned_old_rows]
                selected_from = available or candidates
                if selected_from:
                    old_row = min(selected_from, key=lambda row: (abs(row - new_row), row))
                    new_to_old[new_row] = old_row
                    assigned_old_rows.add(old_row)

            def reuse_category(value):
                text = str(value).strip() if value is not None else ""
                if "新增" in text: return "新增"
                if "复用" in text: return "复用"
                if "利旧" in text: return "利旧"
                return "N/A"

            ws2_final.cell(row=cls.START_ROW - 1, column=cls.PREVIOUS_REUSE_COL).value = "上一轮复用度"
            ws2_final.cell(row=cls.START_ROW - 1, column=cls.PREVIOUS_RESULT_COL).value = "新标号"
            new_results_by_old_row = {}

            # 【修复】：遍历新报告行时，增加空行判断
            for new_row in range(cls.START_ROW, ws2_value.max_row + 1):
                func_val = ws2_value.cell(row=new_row, column=func_col2).value
                desc_val = ws2_value.cell(row=new_row, column=desc_col2).value
                if not func_val and not desc_val:
                    ws2_final.cell(row=new_row, column=cls.PREVIOUS_REUSE_COL).value = None
                    ws2_final.cell(row=new_row, column=cls.PREVIOUS_RESULT_COL).value = None
                    continue

                old_row = new_to_old.get(new_row)
                category = reuse_category(ws1_value.cell(row=old_row, column=12).value) if old_row else "N/A"
                if category == "N/A":
                    new_identifier = "N/A"
                elif category == "新增":
                    new_identifier = 0
                else:
                    remark = ws1_value.cell(row=old_row, column=cls.REMARK_COL).value
                    text = str(remark).strip() if remark is not None else ""
                    if category == "利旧" and not text:
                        new_identifier = "N"
                    elif not text:
                        new_identifier = "N/A"
                    else:
                        match = re.search(r"\d+", text)
                        reference_data = excel1_rows.get(int(match.group())) if match else None
                        target_rows = new_groups.get(group_key(reference_data), []) if reference_data else []
                        new_identifier = (cls._replace_first_number(text, min(target_rows)) if target_rows else text)
                ws2_final.cell(row=new_row, column=cls.PREVIOUS_REUSE_COL).value = category
                ws2_final.cell(row=new_row, column=cls.PREVIOUS_RESULT_COL).value = new_identifier
                if old_row is not None and old_row not in new_results_by_old_row:
                    new_results_by_old_row[old_row] = (new_row, category, new_identifier)

            if progress_callback: progress_callback(60)

            if "Sheet1" in wb1_formula.sheetnames: del wb1_formula["Sheet1"]
            helper = wb1_formula.create_sheet("Sheet1")
            helper.cell(row=3, column=1).value = "旧表行号"
            helper.cell(row=3, column=2).value = "旧表唯一键"
            helper.cell(row=3, column=3).value = "新表行号"
            helper.cell(row=3, column=4).value = "新表唯一键"
            helper.cell(row=3, column=5).value = "新表上一轮复用度"
            helper.cell(row=3, column=6).value = "新表新标号"
            for old_row, data in excel1_rows.items():
                helper.cell(row=old_row, column=1).value = old_row
                helper.cell(row=old_row, column=2).value = data["key"]
                result = new_results_by_old_row.get(old_row)
                if result:
                    new_row, category, new_identifier = result
                    helper.cell(row=old_row, column=3).value = new_row
                    helper.cell(row=old_row, column=4).value = excel2_rows[new_row]["key"]
                    helper.cell(row=old_row, column=5).value = category
                    helper.cell(row=old_row, column=6).value = new_identifier

            ws1_formula.cell(row=cls.START_ROW - 1, column=cls.CURRENT_REUSE_COL).value = "最新复用度"
            ws1_formula.cell(row=cls.START_ROW - 1, column=cls.NEW_ID_COL).value = "新标号"

            # 【修复】：遍历旧报告行写入公式时，增加空行判断
            for old_row in range(cls.START_ROW, ws1_formula.max_row + 1):
                func_val = ws1_formula.cell(row=old_row, column=func_col1).value
                desc_val = ws1_formula.cell(row=old_row, column=desc_col1).value
                if not func_val and not desc_val: continue
                ws1_formula.cell(row=old_row,
                                 column=cls.CURRENT_REUSE_COL).value = f'=IFERROR(VLOOKUP(ROW(),Sheet1!$A:$F,5,FALSE),"N/A")'
                ws1_formula.cell(row=old_row,
                                 column=cls.NEW_ID_COL).value = f'=IFERROR(VLOOKUP(ROW(),Sheet1!$A:$F,6,FALSE),"N/A")'

            if progress_callback: progress_callback(85)
            os.makedirs(output_dir, exist_ok=True)

            # ================= 【核心修复】：使用干净的 project_name 生成文件名 =================
            clean_name = project_name or "重评项目"
            clean_name = re.sub(r"\.xlsx?$", "", clean_name, flags=re.IGNORECASE)
            clean_name = re.sub(r"_?\d{6,14}$", "", clean_name)
            clean_name = re.sub(r"[-_]评估报告$", "", clean_name)
            safe_name = re.sub(r'[\\/*?:"<>|]', "", clean_name).strip("-_ ")
            if not safe_name: safe_name = "重评项目"

            path1 = os.path.join(output_dir, f"{safe_name}_旧报告_对比标注.xlsx")
            path2 = os.path.join(output_dir, f"{safe_name}-评估报告.xlsx")

            wb1_formula.save(path1)
            wb2_final.save(path2)

            if progress_callback: progress_callback(100)
            return path1, path2
        finally:
            wb1_formula.close()
            wb1_value.close()
            wb2_value.close()
            wb2_final.close()