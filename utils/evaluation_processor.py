# utils/evaluation_processor.py
"""
COSMIC 评估处理器
功能:
1. 从功能架构图Word文档提取优化模块
2. 在评估报告Excel的"功能点拆分表"中匹配模块并标注
3. 在"结果计算"工作簿填写送审人天
4. 支持COSMIC规则评估
"""

import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from PySide6.QtCore import QThread, Signal
from datetime import datetime
from difflib import SequenceMatcher
from utils.model_predictor import get_predictor
from utils.runtime_logger import log_error, log_info


class EvaluationWorker(QThread):
    """评估处理工作线程"""

    progress = Signal(int, str)  # 进度值, 消息
    log = Signal(str)  # 日志消息
    finished = Signal(str)  # 完成，返回输出文件路径
    error = Signal(str)  # 错误消息

    # 新增：请求人工验证信号 (结果数据, 回调函数用于返回结果)
    verify_requested = Signal(list, object)

    def __init__(
        self, architecture_doc_path, evaluation_excel_path, architecture_text=None, manday=0, is_manual_mode=False, is_auto_evaluate=False, project_name=None, parent=None
    ):
        super().__init__(parent)
        self.architecture_doc_path = architecture_doc_path
        self.evaluation_excel_path = evaluation_excel_path
        self.architecture_text = architecture_text
        self.manday = manday
        self.is_manual_mode = is_manual_mode
        self.is_auto_evaluate = is_auto_evaluate
        self.project_name = project_name
        self._verified_results = None
        self._verification_done = False

    def run(self):
        """执行评估"""
        try:
            output_path = EvaluationProcessor.process_evaluation(
                architecture_doc_path=self.architecture_doc_path,
                evaluation_excel_path=self.evaluation_excel_path,
                architecture_text=self.architecture_text,
                manday=self.manday,
                is_manual_mode=self.is_manual_mode,
                is_auto_evaluate=self.is_auto_evaluate,
                project_name=self.project_name,
                progress_callback=self._emit_progress,
                log_callback=self._emit_log,
                verification_callback=self._wait_for_verification  # 传入验证回调
            )
            self.finished.emit(output_path)
        except Exception as e:
            import traceback

            traceback.print_exc()
            self.error.emit(str(e))

    def _wait_for_verification(self, results):
        """等待主线程完成人工验证"""
        if not results: return results

        from PySide6.QtCore import QEventLoop
        self._verified_results = results
        self._verification_done = False

        # 发射信号请求验证
        self.verify_requested.emit(results, self._on_verification_finished)

        # 进入局部事件循环等待回复 (在子线程中使用 QEventLoop 是安全的)
        loop = QEventLoop()
        # 我们需要一个方式来打破循环，这里使用一个标志位简单处理
        import time
        while not self._verification_done:
            QThread.msleep(100)
            if self.isInterruptionRequested(): break

        return self._verified_results

    def _on_verification_finished(self, final_results):
        """被主线程回调，保存验证后的结果"""
        self._verified_results = final_results
        self._verification_done = True

    def _emit_progress(self, value, message=""):
        """发送进度信号"""
        self.progress.emit(value, message)

    def _emit_log(self, message):
        """发送日志信号"""
        self.log.emit(message)


class EvaluationProcessor:
    """COSMIC评估处理器"""

    @staticmethod
    def find_column(ws, names, default_col=None, search_rows=15):
        """在工作表中搜索包含特定名称的列"""
        if isinstance(names, str):
            names = [names]

        for r in range(1, search_rows + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val:
                    val_str = str(val).strip()
                    for name in names:
                        if name in val_str:
                            return c
        return default_col

    @staticmethod
    def safe_save(wb, output_path, log_callback=None):
        """安全保存 Workbook，如果权限拒绝则尝试保存为副本"""
        import os
        from datetime import datetime
        try:
            wb.save(output_path)
            if log_callback: log_callback(f"✅ 文件已成功保存: {output_path}")
            return output_path
        except PermissionError:
            if log_callback:
                log_callback(f"⚠️ 无法写入文件 {output_path} (可能被其他程序占用)，尝试保存为副本...")

            base, ext = os.path.splitext(output_path)
            copy_path = f"{base}_副本_{datetime.now().strftime('%H%M%S')}{ext}"
            try:
                wb.save(copy_path)
                if log_callback:
                    log_callback(f"✅ 已成功保存副本: {copy_path}")
                return copy_path
            except Exception as e:
                if log_callback:
                    log_callback(f"❌ 保存副本也失败了: {str(e)}")
                raise
        except Exception as e:
            if log_callback: log_callback(f"❌ 保存文件失败: {str(e)}")
            raise e

    @staticmethod
    def export_preview_excel(source_excel_path, results, log_callback=None):
        """生成预览报告 Excel (匹配 UI 桌面分析表结构，并添加下拉框)"""
        import openpyxl
        from openpyxl.styles import Alignment, PatternFill, Font
        from openpyxl.workbook import Workbook
        from openpyxl.worksheet.datavalidation import DataValidation

        try:
            # 创建全新的预览工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "COSMIC分析预览"

            # 定义表头
            headers = ["行号", "功能过程", "子过程描述", "AI识别结果", "客户原始类型", "判定结果", "智能分析原因"]
            ws.append(headers)

            # 设置样式
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
            center_alignment = Alignment(horizontal='center', vertical='center')

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=c)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_alignment

            # 填入数据
            for res in results:
                row_data = [
                    res.get("row", ""),
                    res.get("group_name", ""),
                    res.get("desc", ""),
                    res.get("move", ""),
                    res.get("orig_move", ""),
                    res.get("row_type", ""),
                    res.get("reason", "")
                ]
                ws.append(row_data)

                # 处理高亮 (AI 建议有改动的行)
                if res.get("is_highlight"):
                    curr_row_idx = ws.max_row
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=curr_row_idx, column=c).fill = yellow_fill

            # 添加下拉框 (数据验证)
            # D 列: AI识别结果 (E, R, X, W) - AI预测不应该返回n
            dv_move = DataValidation(type="list", formula1='"E,R,X,W"', allow_blank=True)
            dv_move.add(f'D2:D{ws.max_row + 100}') # 预留点空间
            ws.add_data_validation(dv_move)

            # F 列: 判定结果 (新增, 复用, 利旧, 优化, 修改, 删除)
            dv_type = DataValidation(type="list", formula1='"新增,复用,利旧,优化,修改,删除"', allow_blank=True)
            dv_type.add(f'F2:F{ws.max_row + 100}')
            ws.add_data_validation(dv_type)

            # 调整列宽
            widths = {'A': 8, 'B': 25, 'C': 50, 'D': 15, 'E': 15, 'F': 15, 'G': 60}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width

            import os
            from datetime import datetime
            base_dir = os.path.dirname(os.path.abspath(source_excel_path))
            preview_path = os.path.join(base_dir, f"PREVIEW_评估分析_{datetime.now().strftime('%H%M%S')}.xlsx")
            wb.save(preview_path)
            return preview_path
        except Exception as e:
            if log_callback: log_callback(f"导出分析预览失败: {str(e)}")
            return None

    @staticmethod
    def load_results_from_excel(preview_excel_path, current_results, log_callback=None):
        """从专用的预览分析表读取修改后的结果"""
        import openpyxl
        try:
            wb = openpyxl.load_workbook(preview_excel_path, data_only=True)
            ws = wb.active

            # 列定义: A:行号, D:动作, F:判定结果
            updated_count = 0
            new_results = [res.copy() for res in current_results]
            results_map = {res["row"]: i for i, res in enumerate(new_results)}

            # 遍历数据行 (从第二行开始，跳过表头)
            for r_idx in range(2, ws.max_row + 1):
                row_no_val = ws.cell(row=r_idx, column=1).value
                if row_no_val is None: continue

                # 跳过表头行（有时表头会出现在数据区域）
                if str(row_no_val).strip() == "行号":
                    continue

                try:
                    row_no = int(row_no_val)
                    if row_no in results_map:
                        target_res = new_results[results_map[row_no]]

                        # 读取 AI 修改项
                        new_move = str(ws.cell(row=r_idx, column=4).value or "").strip().upper()
                        new_rt = str(ws.cell(row=r_idx, column=6).value or "").strip()

                        if new_move in ["E", "R", "X", "W", "N"]:
                            m = "n" if new_move == "N" else new_move
                            if m != str(target_res.get("move", "")).upper():
                                target_res["move"] = m
                                updated_count += 1

                        if new_rt in ["新增", "复用", "利旧", "优化"]:
                            if new_rt != target_res.get("row_type", ""):
                                target_res["row_type"] = new_rt
                                updated_count += 1
                except (ValueError, TypeError):
                    continue

            if log_callback: log_callback(f"从 Excel 中成功同步了 {updated_count} 处改动")
            return new_results
        except Exception as e:
            if log_callback: log_callback(f"同步预览 Excel 失败: {str(e)}")
            return current_results

    @staticmethod
    def parse_manual_input(text, progress_callback=None, log_callback=None):
        """
        解析手动输入的优化模块列表
        返回: [(模块编号, 模块名称, 级别, 路径), ...]
        """
        if log_callback:
            log_callback("正在解析手动输入的优化模块...")

        optimization_modules = []
        lines = text.strip().split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 匹配格式: 编号 模块名称
            match = re.match(r'([\d\.]+)\s+(.+)', line)
            if match:
                num = match.group(1).strip('.')
                title = match.group(2).strip()
                # 根据编号判断级别
                level = num.count('.') + 1
                optimization_modules.append((num, title, level, ""))
            else:
                # 如果没有编号，则将整行作为名称，默认为二级
                # 这是为了兼容“或直接输入名称”的情况
                optimization_modules.append(("-", line, 2, ""))

            if log_callback:
                if match:
                    log_callback(f"解析到模块: {match.group(1).strip('.')} {match.group(2).strip()}")
                else:
                    log_callback(f"解析到名称 (默认二级): {line}")

        if progress_callback:
            progress_callback(30, f"解析到 {len(optimization_modules)} 个优化模块")

        if log_callback:
            log_callback(f"共解析到 {len(optimization_modules)} 个优化模块")

        return optimization_modules

    # @staticmethod
    # def infer_group_results(group, proc_type, essential_moves, trust_move=False, proc_type_source="rule"):
    #     """
    #         Args:
    #             proc_type_source: "ai" 或 "rule"，标识 proc_type 的来源
    #     """
    #     """核心推理逻辑：对单一功能过程中的所有行进行判定"""
    #     name = group.get("name", "").lower()
    #
    #     # 加载规则关键字来自配置
    #     from extend.matcher_config import MatcherConfig
    #     config = MatcherConfig.load()
    #     evaluation_rules = config.get("evaluation_rules", {})
    #     legacy_kws = evaluation_rules.get("legacy_keywords", ["利旧", "内存", "缓存", "校验", "由于"])
    #     reuse_kws = evaluation_rules.get("reuse_keywords", ["复用", "继承", "沿用"])
    #
    #     def run_eval_internal(current_proc_type, current_essential_moves):
    #         current_results = []
    #         seen_in_group = set()
    #
    #         # --- 第一遍：分配确定性最高的动作 ---
    #         temp_rows_data = []
    #         for idx, data in enumerate(group["rows"]):
    #             r = data["row"]
    #             move_char = data.get("move", "n")
    #             ai_move = data.get("ai_move", "n")
    #             all_ai_types = data.get("all_ai_types", ai_move)
    #             excel_move = data.get("excel_move", "")
    #             is_o = data.get("is_optimization", False)
    #             desc = data["desc"]
    #             matched_kws = data.get("matched_kws", [])
    #             matched_categories = data.get("matched_categories", [])
    #             detail_reason = data.get("detail_reason_pre", "")
    #             is_real_leiju = data.get("is_real_leiju", False)
    #
    #             final_move_decided = "n"
    #             thinking_lines = []
    #             is_highlight = False
    #
    #             if trust_move:
    #                 final_move_decided = move_char
    #                 thinking_lines = [f"          思考过程：[人工校验模式] 信任用户选择的动作：{move_char}"]
    #             else:
    #                 kws_str = "/".join(matched_kws) if matched_kws else "未找到"
    #                 user_intent = excel_move if excel_move else "未填"
    #
    #                 # 修改逻辑：只有真正的利旧才标记为n
    #                 if is_real_leiju:
    #                     thinking_lines.append(f"          思考过程：根据关键词：[{kws_str}]，识别到利旧特征，判定为利旧(n)")
    #                     final_move_decided = "n"
    #                 elif not matched_categories:
    #                     if excel_move and excel_move != "n":
    #                         thinking_lines.append(f"          思考过程：根据关键词：[未找到]，模型无法确认动作，尊重用户填写结果：{excel_move}")
    #                         final_move_decided = excel_move
    #                         is_highlight = True
    #                     else:
    #                         thinking_lines.append(f"          思考过程：根据关键词：[未找到]，模型与用户均未确认动作，默认为利旧(n)")
    #                         final_move_decided = "n"
    #                 else:
    #                     # 核心改进：智能补全逻辑
    #                     # 如果有多个匹配动作，优先选择当前过程“还缺少”的必要动作
    #                     best_move = ai_move
    #                     if len(matched_categories) > 1:
    #                         # 过滤出当前过程需要的动作
    #                         potential_essentials = [m for m in matched_categories if m in current_essential_moves]
    #                         if potential_essentials:
    #                             # 如果 ai_move 已经在组内看过了，或者 ai_move 不是必要动作，尝试换一个
    #                             remaining_essentials = [m for m in potential_essentials if m not in seen_in_group]
    #                             if remaining_essentials:
    #                                 best_move = remaining_essentials[0]
    #                                 if best_move != ai_move:
    #                                     thinking_lines.append(f"          思考过程：检测到多重动作 {all_ai_types}，由于过程缺少 {best_move}，系统智能切换识别结果为 {best_move}")
    #
    #                     if excel_move and excel_move in matched_categories:
    #                         thinking_lines.append(f"          思考过程：根据关键词：[{kws_str}]，我认为动作是{all_ai_types}，用户认为是{excel_move}，最终结果应该是{excel_move}")
    #                         final_move_decided = excel_move
    #                     elif excel_move and best_move != excel_move:
    #                         thinking_lines.append(f"          思考过程：根据关键词：[{kws_str}]，我认为动作是{all_ai_types}，用户认为是{excel_move}，应该是用户填写错误，最终结果应该是{best_move}")
    #                         final_move_decided = best_move
    #                         is_highlight = True
    #                     else:
    #                         thinking_lines.append(f"          思考过程：根据关键词：[{kws_str}]，我认为动作是{all_ai_types}，最终结果是{best_move}")
    #                         final_move_decided = best_move
    #                         if excel_move and best_move != excel_move:
    #                             is_highlight = True
    #
    #             # 记录以便分配
    #             temp_rows_data.append({
    #                 "r": r, "final_move": final_move_decided, "is_o": is_o, "desc": desc,
    #                 "thinking": thinking_lines, "is_highlight": is_highlight, "ai_move": ai_move
    #             })
    #             # 更新组内已见动作
    #             if final_move_decided in current_essential_moves:
    #                 seen_in_group.add(final_move_decided)
    #
    #         # --- 第二遍：构造结果 ---
    #         seen_final = set()
    #         for item in temp_rows_data:
    #             m = item["final_move"]
    #             if m == "n" or m not in current_essential_moves:
    #                 current_results.append({
    #                     "row": item["r"], "value": "n",
    #                     "reason": f"利旧(n)-与{current_proc_type}动作不符" if m != "n" else "利旧(n)-未识别到有效动作",
    #                     "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
    #                     "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
    #                 })
    #                 continue
    #
    #             if m in seen_final:
    #                 current_results.append({
    #                     "row": item["r"], "value": "n",
    #                     "reason": f"忽略(n)-动作{m}已计入",
    #                     "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
    #                     "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
    #                 })
    #                 continue
    #
    #             seen_final.add(m)
    #             current_results.append({
    #                 "row": item["r"], "value": "PENDING", "reason": "",
    #                 "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
    #                 "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
    #             })
    #
    #         found = {res["move"] for res in current_results if res["value"] != "n"}
    #         if current_proc_type == "EW":
    #             result_is_comp = ("E" in found) and ("W" in found)
    #         elif current_proc_type == "EX":
    #             result_is_comp = ("E" in found) and ("X" in found)
    #         elif current_proc_type == "EXEW":
    #             # EXEW: 2个E、1个X、1个W (至少需要E、X、W各1个)
    #             result_is_comp = ("E" in found) and ("X" in found) and ("W" in found)
    #         elif current_proc_type == "EXEX":
    #             # EXEX: 2个E、2个X (至少需要E、X各1个)
    #             result_is_comp = ("E" in found) and ("X" in found)
    #         else:
    #             # ERX 必须包含 E, R, X
    #             result_is_comp = ("E" in found) and ("R" in found) and ("X" in found)
    #
    #         return current_results, result_is_comp, found
    #
    #     # 初始评估
    #     group_results, is_complete_res, found_moves_res = run_eval_internal(proc_type, essential_moves)
    #
    #     # 自动切换类型逻辑
    #     final_proc_type = proc_type
    #     final_essential_moves = essential_moves
    #
    #     force_stay_ew = (proc_type == "EW" and any(kw in name for kw in ["新增", "删除", "修改", "保存", "编辑", "更新", "注销"]))
    #     if not is_complete_res and not force_stay_ew:
    #         # 1. 尝试切换主类型 (EW <-> ERX)
    #         alt_type = "ERX" if proc_type == "EW" else "EW"
    #         alt_essential = ["E", "R", "X"] if alt_type == "ERX" else ["E", "W"]
    #         alt_results, alt_is_complete, alt_found = run_eval_internal(alt_type, alt_essential)
    #         if alt_is_complete:
    #             final_proc_type, final_essential_moves, group_results, is_complete_res, found_moves_res = \
    #                 alt_type, alt_essential, alt_results, alt_is_complete, alt_found
    #
    #         # 2. 如果依然不完整，尝试降级判定 (ERX -> EX)
    #         if not is_complete_res and (proc_type == "ERX" or alt_type == "ERX"):
    #             ex_results, ex_is_complete, ex_found = run_eval_internal("EX", ["E", "X"])
    #             if ex_is_complete:
    #                 final_proc_type, final_essential_moves, group_results, is_complete_res, found_moves_res = \
    #                     "EX", ["E", "X"], ex_results, True, ex_found
    #
    #     # 后处理：设置最终结果
    #     summary_thinking = ""
    #     found_chars = "".join(sorted(list(found_moves_res)))
    #     # 如果 proc_type 来自 AI 预测，优先显示"AI 预测"
    #     if proc_type_source == "ai":
    #         trigger_kw = "🤖AI 预测"
    #     else:
    #         trigger_kw = "未找到"
    #         for kw in (evaluation_rules.get("ew_group_keywords", []) + evaluation_rules.get("erx_group_keywords", [])):
    #             if kw in name:
    #                 trigger_kw = kw
    #                 break
    #
    #     if not is_complete_res:
    #         missing = [m for m in final_essential_moves if m not in found_moves_res]
    #         for res in group_results:
    #             res["value"] = "n"
    #             res["reason"] = f"利旧(n)-过程不完整({final_proc_type}缺少必要动作: {','.join(missing)})"
    #
    #         actual_found = found_chars if found_chars else "无"
    #         summary_thinking = f"{group['rows'][0]['row']}-{group['rows'][-1]['row']}行功能过程思考过程：识别关键词：[{trigger_kw}]，判定类型: {final_proc_type}，但是缺少必要动作: {','.join(missing)}，最终结果应该是{actual_found}"
    #     else:
    #         for res in group_results:
    #             if res["value"] == "n": continue
    #             m_char, is_o = res["move"], res["is_o"]
    #             if is_o:
    #                 res["value"], res["reason"] = "O", "优化(O)"
    #             else:
    #                 res["value"], res["reason"] = None, f"新增 [首个{m_char}]"
    #         summary_thinking = f"{group['rows'][0]['row']}-{group['rows'][-1]['row']}行功能过程思考过程：识别关键词：[{trigger_kw}]，最后的结果是{final_proc_type}，判定类型符合预期，最终结果是{found_chars}"
    #
    #     for res in group_results:
    #         res["summary_thinking"] = summary_thinking
    #         res["final_proc_type"] = final_proc_type
    #
    #     return group_results

    @staticmethod
    def infer_group_results(group, proc_type, essential_moves, trust_move=False, proc_type_source="rule"):
        """
        核心推理逻辑：对单一功能过程中的所有行进行判定
        Args:
            proc_type_source: "ai" 或 "rule"，标识 proc_type 的来源
        """
        name = group.get("name", "").lower()

        # 加载规则关键字来自配置
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        evaluation_rules = config.get("evaluation_rules", {})
        legacy_kws = evaluation_rules.get("legacy_keywords", ["利旧", "内存", "缓存", "校验", "由于"])
        reuse_kws = evaluation_rules.get("reuse_keywords", ["复用", "继承", "沿用"])

        def run_eval_internal(current_proc_type, current_essential_moves):
            current_results = []
            seen_in_group = set()

            # --- 第一遍：分配确定性最高的动作 ---
            temp_rows_data = []
            for idx, data in enumerate(group["rows"]):
                r = data["row"]
                move_char = data.get("move", "n")
                ai_move = data.get("ai_move", "n")
                all_ai_types = data.get("all_ai_types", ai_move)
                excel_move = data.get("excel_move", "")
                is_o = data.get("is_optimization", False)
                desc = data["desc"]
                matched_kws = data.get("matched_kws", [])
                matched_categories = data.get("matched_categories", [])
                detail_reason = data.get("detail_reason_pre", "")
                is_real_leiju = data.get("is_real_leiju", False)

                final_move_decided = "n"
                thinking_lines = []
                is_highlight = False

                if trust_move:
                    final_move_decided = move_char
                    thinking_lines = [f"          思考过程：[人工校验模式] 信任用户选择的动作：{move_char}"]
                else:
                    kws_str = "/".join(matched_kws) if matched_kws else "未找到"
                    user_intent = excel_move if excel_move else "未填"

                    # 修改逻辑：只有真正的利旧才标记为 n
                    if is_real_leiju:
                        thinking_lines.append(
                            f"          思考过程：根据关键词：[{kws_str}]，识别到利旧特征，判定为利旧 (n)")
                        final_move_decided = "n"
                    elif not matched_categories:
                        if excel_move and excel_move != "n":
                            thinking_lines.append(
                                f"          思考过程：根据关键词：[未找到]，模型无法确认动作，尊重用户填写结果：{excel_move}")
                            final_move_decided = excel_move
                            is_highlight = True
                        else:
                            thinking_lines.append(
                                f"          思考过程：根据关键词：[未找到]，模型与用户均未确认动作，默认为利旧 (n)")
                            final_move_decided = "n"
                    else:
                        # 核心改进：智能补全逻辑 + 用户优先策略
                        best_move = ai_move

                        # 🔧 关键改进：如果同时有 W 和 X，且描述含"结果"或"返回"，优先判定为 X (出口)
                        if "W" in matched_categories and "X" in matched_categories:
                            if any(kw in desc.lower() for kw in ["结果", "返回", "响应", "输出"]):
                                best_move = "X"
                                thinking_lines.append(f"          思考过程：检测到多重动作 [W/X]，由于描述含“结果/返回”，识别结果为X")

                        # 🔧 关键改进：如果有多个匹配动作，优先选择当前过程"还缺少"的必要动作
                        if len(matched_categories) > 1 and best_move not in ["X"]: # 如果已经确定是 X 且有明确特征，不再切换到 W
                            # 过滤出当前过程需要的动作
                            potential_essentials = [m for m in matched_categories if m in current_essential_moves]
                            if potential_essentials:
                                # 如果 ai_move 已经在组内看过了，或者 ai_move 不是必要动作，尝试换一个
                                remaining_essentials = [m for m in potential_essentials if m not in seen_in_group]
                                if remaining_essentials:
                                    best_move = remaining_essentials[0]
                                    if best_move != ai_move:
                                        thinking_lines.append(
                                            f"          思考过程：检测到多重动作 [{all_ai_types}]，由于过程缺少{best_move}，系统智能切换识别结果为{best_move}")

                        # 🔧 关键改进：用户优先策略 - 如果用户填写的在匹配结果中，优先采用
                        if excel_move and excel_move in matched_categories:
                            thinking_lines.append(
                                f"          思考过程：检测到多重动作 [{all_ai_types}]，用户填写为{excel_move}，优先采用用户值")
                            final_move_decided = excel_move
                        elif excel_move and best_move != excel_move:
                            thinking_lines.append(
                                f"          思考过程：检测到多重动作 [{all_ai_types}]，用户认为是{excel_move}，应该是用户填写错误，最终结果应该是{best_move}")
                            final_move_decided = best_move
                            is_highlight = True
                        else:
                            thinking_lines.append(
                                f"          思考过程：检测到多重动作 [{all_ai_types}]，最终结果是{best_move}")
                            final_move_decided = best_move
                            if excel_move and best_move != excel_move:
                                is_highlight = True

                # 记录以便分配
                temp_rows_data.append({
                    "r": r, "final_move": final_move_decided, "is_o": is_o, "desc": desc,
                    "thinking": thinking_lines, "is_highlight": is_highlight, "ai_move": ai_move
                })
                # 更新组内已见动作
                if final_move_decided in current_essential_moves:
                    seen_in_group.add(final_move_decided)

            # --- 第二遍：构造结果 ---
            seen_final = set()
            for item in temp_rows_data:
                m = item["final_move"]

                # 🔧 修复：只有真正的 n 才标记为利旧
                if m == "n":
                    current_results.append({
                        "row": item["r"], "value": "n",
                        "reason": "利旧 (n)-未识别到有效动作",
                        "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
                        "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
                    })
                    continue

                # 🔧 修复：动作不在必备动作中，不强制标记为利旧，而是记录为 PENDING
                if m not in current_essential_moves:
                    # 🔧 关键改进：EW 过程中的 X 动作 (如"返回结果")，根据用户要求标记为忽略 (n)
                    is_crud_ew = any(kw in name.lower() for kw in ["新增", "删除", "修改", "保存", "编辑", "更新", "注销"])
                    if current_proc_type == "EW" and m == "X" and is_crud_ew:
                        thinking_lines = item["thinking"]
                        # 如果描述包含"结果"或"返回"，且被识别为 X，则添加特定思考过程
                        thinking_lines.append(f"          思考过程：检测到多重动作 [W/X]，用户认为是X，最终结果应该是X")

                        current_results.append({
                            "row": item["r"], "value": "n",
                            "reason": f"忽略 (n)-动作{m}已计入",
                            "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
                            "thinking": "\n".join(thinking_lines), "is_highlight": item["is_highlight"]
                        })
                        continue

                    current_results.append({
                        "row": item["r"], "value": "PENDING", "reason": "",
                        "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
                        "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
                    })
                    continue

                if m in seen_final:
                    current_results.append({
                        "row": item["r"], "value": "n",
                        "reason": f"忽略 (n)-动作{m}已计入",
                        "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
                        "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
                    })
                    continue

                seen_final.add(m)
                current_results.append({
                    "row": item["r"], "value": "PENDING", "reason": "",
                    "move": m, "ai_move": item["ai_move"], "is_o": item["is_o"], "desc": item["desc"],
                    "thinking": "\n".join(item["thinking"]), "is_highlight": item["is_highlight"]
                })

            # 🔧 修复：收集所有非利旧的动作（包括 PENDING）
            found = {res["move"] for res in current_results if res["value"] != "n" and res["move"] != "n"}

            # 🔧 修复：如果有 X 动作但过程类型是 EW，标记为不完整以触发类型切换
            has_x_action = "X" in found
            if current_proc_type == "EW":
                result_is_comp = ("E" in found) and ("W" in found)
                # 如果有 X 动作，强制标记为不完整，触发类型切换
                if has_x_action:
                    result_is_comp = False
            elif current_proc_type == "EX":
                result_is_comp = ("E" in found) and ("X" in found)
            elif current_proc_type == "EXEW":
                result_is_comp = ("E" in found) and ("X" in found) and ("W" in found)
            elif current_proc_type == "EXEX":
                result_is_comp = ("E" in found) and ("X" in found)
            else:
                # ERX 必须包含 E, R, X
                result_is_comp = ("E" in found) and ("R" in found) and ("X" in found)

            return current_results, result_is_comp, found

        # 初始评估
        group_results, is_complete_res, found_moves_res = run_eval_internal(proc_type, essential_moves)

        # 自动切换类型逻辑
        final_proc_type = proc_type
        final_essential_moves = essential_moves

        force_stay_ew = (proc_type == "EW" and any(
            kw in name for kw in ["新增", "删除", "修改", "保存", "编辑", "更新", "注销"]))
        if not is_complete_res and not force_stay_ew:
            # 1. 尝试切换主类型 (EW <-> ERX)
            alt_type = "ERX" if proc_type == "EW" else "EW"
            alt_essential = ["E", "R", "X"] if alt_type == "ERX" else ["E", "W"]
            alt_results, alt_is_complete, alt_found = run_eval_internal(alt_type, alt_essential)
            if alt_is_complete:
                final_proc_type, final_essential_moves, group_results, is_complete_res, found_moves_res = \
                    alt_type, alt_essential, alt_results, alt_is_complete, alt_found

            # 2. 如果依然不完整，尝试降级判定 (ERX -> EX)
            if not is_complete_res and (proc_type == "ERX" or alt_type == "ERX"):
                ex_results, ex_is_complete, ex_found = run_eval_internal("EX", ["E", "X"])
                if ex_is_complete:
                    final_proc_type, final_essential_moves, group_results, is_complete_res, found_moves_res = \
                        "EX", ["E", "X"], ex_results, True, ex_found

        # 后处理：设置最终结果
        summary_thinking = ""
        found_chars = " ".join(sorted(list(found_moves_res)))

        # 🔧 关键改进：如果 proc_type 来自 AI 预测，优先显示"AI 预测"
        if proc_type_source == "ai":
            trigger_kw = "🤖AI 预测"
        else:
            trigger_kw = "未找到"
            for kw in (evaluation_rules.get("ew_group_keywords", []) + evaluation_rules.get("erx_group_keywords", [])):
                if kw in name:
                    trigger_kw = kw
                    break

        if not is_complete_res:
            missing = [m for m in final_essential_moves if m not in found_moves_res]

            # 🔧 修复：检查是否有额外动作（如 EW 过程但有 X 动作）
            extra_moves = found_moves_res - set(final_essential_moves)

            if extra_moves and "X" in extra_moves and final_proc_type == "EW":
                # 切换到 ERX 类型
                final_proc_type = "ERX"
                final_essential_moves = ["E", "R", "X"]
                # 重新标记结果
                for res in group_results:
                    if res["value"] == "n" and res["move"] in ["E", "R", "X"]:
                        res["value"] = None  # 新增
                        res["reason"] = f"新增 [首个{res['move']}]"

            # 如果依然不完整，再标记为利旧
            missing = [m for m in final_essential_moves if m not in found_moves_res]
            if missing:
                for res in group_results:
                    if res["value"] == "PENDING":
                        if res["move"] in final_essential_moves:
                            res["value"] = None
                            res["reason"] = f"新增 [首个{res['move']}]"
                        else:
                            res["value"] = "n"
                            res["reason"] = f"利旧 (n)-与{final_proc_type}动作不符"

            actual_found = found_chars if found_chars else "无"
            summary_thinking = f"{group['rows'][0]['row']}-{group['rows'][-1]['row']}行功能过程思考过程：识别关键词：[{trigger_kw}]，判定类型：{final_proc_type}，但是缺少必要动作：{','.join(missing)}，最终结果应该是{actual_found}"
        else:
            for res in group_results:
                if res["value"] == "n": continue
                m_char, is_o = res["move"], res["is_o"]
                if is_o:
                    res["value"], res["reason"] = "O", "优化 (O)"
                elif res["value"] == "PENDING":
                    res["value"], res["reason"] = None, f"新增 [首个{m_char}]"
            summary_thinking = f"{group['rows'][0]['row']}-{group['rows'][-1]['row']}行功能过程思考过程：识别关键词：[{trigger_kw}]，最后的结果是{final_proc_type}，判定类型符合预期，最终结果是{found_chars}"

        for res in group_results:
            res["summary_thinking"] = summary_thinking
            res["final_proc_type"] = final_proc_type

        return group_results

    @staticmethod
    def re_evaluate_inference(all_results):
        """人工修改 move 后重新执行推理循环 (仅限当前所有行)"""
        # 加载规则关键字
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        rules = config.get("evaluation_rules", {})
        ew_keywords = rules.get("ew_group_keywords", ["新增", "删除", "修改", "导入", "上传", "创建", "保存", "更新", "同步", "处理", "注销", "重置", "通过", "拒绝", "配置", "变更"])
        erx_keywords = rules.get("erx_group_keywords", ["查看", "查询", "下载", "导出", "搜索", "浏览", "列表", "获取", "详情", "展示", "显示", "统计", "分析", "报表", "图表"])
        reuse_threshold = config.get("reuse_threshold", 70)

        # 1. 按照 group_name 重新分组
        groups_dict = {}
        # 为了保证推理的上下文完整性，我们需要按顺序处理
        for res in all_results:
            g_name = res.get("group_name", "Unknown")
            if g_name not in groups_dict:
                groups_dict[g_name] = {"name": g_name, "l3": res.get("l3", ""), "rows": []}

            # 构造原始数据格式，确保使用当前 UI 上的 move
            # 重新提取关键字信息，以便 infer 逻辑正常工作 (针对智能补全)
            desc_val = res.get("desc", "")
            h_text = desc_val.strip().lower()
            import re
            h_text = re.sub(r"^[0-9.\-\s]+", "", h_text)

            entry_kws = rules.get("entry_keywords", ["接收", "送入", "输入", "点击", "发起", "上传", "跳转"])
            read_kws = rules.get("read_keywords", ["读取", "检索", "查询", "获取", "通过查询"])
            exit_kws = rules.get("exit_keywords", ["返回", "响应", "输出", "导出", "下载", "推送", "回显", "显示", "展示", "提示", "结果"])
            write_kws = rules.get("write_keywords", ["写入", "更新", "保存", "存储", "修改", "逻辑处理", "删除", "新增", "创建", "同步", "注销", "移除", "清除"])

            matched_categories = []
            matched_kws = []
            for kw in entry_kws:
                if kw in h_text: matched_categories.append("E"); matched_kws.append(kw)
            for kw in read_kws:
                if kw in h_text: matched_categories.append("R"); matched_kws.append(kw)
            for kw in exit_kws:
                if kw in h_text: matched_categories.append("X"); matched_kws.append(kw)
            for kw in write_kws:
                if kw in h_text: matched_categories.append("W"); matched_kws.append(kw)

            # 智能补全初始类别
            if not matched_categories:
                for kw in ["请求", "提交", "选择", "勾选"]:
                    if kw in h_text: matched_categories.append("E"); matched_kws.append(kw)

            matched_categories = list(dict.fromkeys(matched_categories))
            all_ai_types = "/".join(matched_categories) if matched_categories else "n"

            groups_dict[g_name]["rows"].append({
                "row": res["row"],
                "desc": desc_val,
                "move": res.get("move", "n"),
                "ai_move": res.get("ai_move", ""), # 增加 ai_move 传递
                "all_ai_types": all_ai_types,
                "matched_categories": matched_categories,
                "matched_kws": matched_kws,
                "excel_move": res.get("orig_move", ""),
                "is_optimization": (res.get("row_type") == "优化"),
                "detail_reason_pre": "" # 重置以便 infer 重新扫描
            })

        # 2. 对每个组重新推理 (模仿 auto_evaluate_cosmic 的两阶段模式)
        final_results = []
        seen_groups = []
        for r in all_results:
            gn = r.get("group_name", "Unknown")
            if gn not in seen_groups: seen_groups.append(gn)

        # 阶段一：独立分析
        independent_analyses = []
        for g_name in seen_groups:
            group = groups_dict[g_name]
            # 重新判断 proc_type (根据用户修改后的 moves)
            name_lower = g_name.lower()
            # 优先判定 EW
            is_ew_kw = any(kw in name_lower for kw in ew_keywords)
            is_erx_kw = any(kw in name_lower for kw in erx_keywords)

            # 这里的逻辑应当尊重过程名，但也要看 moves 是否包含 W
            moves_in_group = "".join([r["move"] for r in group["rows"]])
            if is_ew_kw or "W" in moves_in_group:
                proc_type, essential_moves = "EW", ["E", "W"]
            elif is_erx_kw:
                proc_type, essential_moves = "ERX", ["E", "R", "X"]
            else:
                proc_type, essential_moves = "ERX", ["E", "R", "X"]

            # 推理：infer_group_results 现在完全信任传入的 move
            # 传入的 rows 已经带有重新提取的 matched_categories
            group_results = EvaluationProcessor.infer_group_results(group, proc_type, essential_moves, trust_move=True)
            actual_type = group_results[0].get("final_proc_type", proc_type) if group_results else proc_type

            independent_analyses.append({
                "name": g_name,
                "results": group_results,
                "type": actual_type
            })

        # 阶段二：复用判定
        global_process_benchmarks = []
        for analysis in independent_analyses:
            g_name = analysis["name"]
            group_results = analysis["results"]
            actual_type = analysis["type"]

            best_group_match = None
            max_group_sim = 0
            for prev_group in global_process_benchmarks:
                type_compatible = False
                prev_type = prev_group["type"]
                if actual_type == "EW" and prev_type == "EW":
                    type_compatible = True
                elif actual_type in ["ERX", "EX"] and prev_type in ["ERX", "EX"]:
                    type_compatible = True

                if type_compatible:
                    sim = EvaluationProcessor.calculate_similarity(g_name, prev_group["name"])
                    if sim > max_group_sim and sim >= reuse_threshold:
                        max_group_sim = sim
                        best_group_match = prev_group

            if best_group_match:
                new_group_results = []
                for orig_res in group_results:
                    r, m = orig_res["row"], orig_res["move"]
                    is_o = orig_res.get("is_o", False)
                    desc = orig_res.get("desc", "")

                    if orig_res["value"] == "n":
                        val, reason = "n", orig_res["reason"]
                    else:
                        move_match = next((ref for ref in best_group_match["rows"] if ref["move"] == m and ref["value"] != "n"), None)
                        if move_match:
                            ref_val = move_match["value"]
                            target_ref = move_match["row"] if (ref_val == "O" or ref_val is None) else ref_val
                            val = "O" if is_o else int(target_ref)
                            reason = f"{'优化(O)' if is_o else '复用'}({target_ref}) <- [{best_group_match['name']}]"
                        else:
                            val, reason = orig_res["value"], orig_res["reason"]

                    new_group_results.append({
                        "row": r, "value": val, "reason": reason, "move": m, "desc": desc,
                        "is_highlight": orig_res.get("is_highlight", False),
                        "thinking": f"          思考过程：[复用模式] 系统匹配到参考过程 {best_group_match['name']}，本行处理逻辑：{reason}"
                    })
                group_results = new_group_results

            # 回填基础字段和原始数据
            for r in group_results:
                r["group_name"] = g_name
                orig_data = next((item for item in all_results if item["row"] == r["row"]), {})
                r["orig_move"] = orig_data.get("orig_move", "")
                r["orig_l"] = orig_data.get("orig_l", "")
                r["l3"] = orig_data.get("l3", "")
                r["ai_move"] = orig_data.get("ai_move", "")

                # 转换 row_type
                v = r.get("value")
                if v == "n": r["row_type"] = "利旧"
                elif v == "O": r["row_type"] = "优化"
                elif v is None: r["row_type"] = "新增"
                else: r["row_type"] = "复用"

                final_results.append(r)

            # 加入基准池供后续组使用
            global_process_benchmarks.append({"name": g_name, "rows": group_results, "type": actual_type})

        return final_results

    @staticmethod
    def calculate_similarity(s1, s2):
        """计算相似度 (0-100)，整合 SequenceMatcher 与 业务关键字逻辑"""
        if not s1 or not s2: return 0
        from difflib import SequenceMatcher

        def get_clean_val(v):
            if v is None: return ""
            # 去除符号
            return re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]", "", str(v)).lower()

        s1_clean, s2_clean = get_clean_val(s1), get_clean_val(s2)
        if s1_clean == s2_clean: return 100

        # 针对 COSMIC 业务场景的优先级加权：
        # 如果一方包含另一方，给予高分
        if s1_clean and s2_clean and (s1_clean in s2_clean or s2_clean in s1_clean):
            return 95

        # 针对“新增/修改/查看”等动作的特殊处理
        # 提取核心业务词：去掉常见的 COSMIC 动作关键词
        keywords = ["新增", "修改", "编辑", "删除", "批量", "保存", "更新", "查看", "查询", "列表", "详情", "统计", "分析", "展示"]
        s1_bus = s1_clean
        s2_bus = s2_clean
        for kw in keywords:
            s1_bus = s1_bus.replace(kw, "")
            s2_bus = s2_bus.replace(kw, "")

        # 如果去掉动作关键词后完全一致，说明是同一个业务的不同操作，复用概率极高
        if s1_bus and s2_bus and s1_bus == s2_bus:
            return 98

        # 如果业务词高度相似（如 酒店维护信息 vs 酒店信息）
        if s1_bus and s2_bus and (s1_bus in s2_bus or s2_bus in s1_bus):
            return 92

        # 使用 SequenceMatcher 进行最终兜底
        return SequenceMatcher(None, s1_clean, s2_clean).ratio() * 100

    @staticmethod
    def find_sheet_by_names(wb, possible_names):
        """根据可能的名称查找工作表"""
        for name in possible_names:
            if name in wb.sheetnames:
                return wb[name]
        # 模糊匹配
        for sheet_name in wb.sheetnames:
            for name in possible_names:
                if name in sheet_name:
                    return wb[sheet_name]
        return None

    @staticmethod
    def match_and_mark_modules(
        ws_split,
        optimization_modules,
        start_row=5,
        progress_callback=None,
        log_callback=None,
    ):
        """
        在功能点拆分表中匹配优化模块并在 P 列标注'O'
        策略：寻找最优匹配行，避免过度标注
        """
        log_info('\n' + '=' * 50)
        log_info('🔍 启动优化模块精准匹配流程...')

        level1_col = EvaluationProcessor.find_column(ws_split, ["一级模块", "一级", "L1"], default_col=2)
        level2_col = EvaluationProcessor.find_column(ws_split, ["二级模块", "二级", "L2"], default_col=3)
        level3_col = EvaluationProcessor.find_column(ws_split, ["三级模块", "三级", "L3"], default_col=4)
        remark_col = EvaluationProcessor.find_column(ws_split, ["备注", "Remark", "说明"], default_col=16)
        orange_fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

        # 1. 预处理：去重并提取纯净名称
        unique_modules = {}
        for num, title, level, path in optimization_modules:
            # 移除编号前缀和括号内容
            clean_title = re.sub(r"^\d+[\.\s]*", "", title).strip()
            clean_title = re.sub(r"[（(].*?[)）]", "", clean_title).strip()
            if clean_title:
                unique_modules[clean_title] = {"num": num, "level": level}

        log_info(f'待匹配优化模块 (共 {len(unique_modules)} 个): {list(unique_modules.keys())}')

        # 2. 遍历 Excel 寻找每个模块的最佳匹配行
        matched_rows = set()
        for mod_title, info in unique_modules.items():
            best_score = 0
            best_row = -1
            best_col = -1

            # 确定优先搜索列 (兼容字符串 "二级" 或 整数 2)
            target_level = str(info.get("level", ""))
            if "1" in target_level or "一级" in target_level:
                search_order = [level1_col, level2_col, level3_col]
            elif "2" in target_level or "二级" in target_level:
                search_order = [level2_col, level3_col, level1_col]
            else:
                search_order = [level3_col, level2_col, level1_col]

            # 在三级模块、二级模块和一级模块列中寻找最佳匹配
            for r in range(start_row, ws_split.max_row + 1):
                for col in search_order:
                    val = ws_split.cell(row=r, column=col).value
                    if not val: continue

                    score = EvaluationProcessor.calculate_similarity(mod_title, str(val))

                    # 权重调整：如果列匹配用户选择的级别，加分
                    if col == search_order[0]:
                        score += 5

                    if score > best_score:
                        best_score = score
                        best_row = r
                        best_col = col

            # 阈值判定：75分以上视为找到 (扣除权重后的真实分数需过 70)
            if best_score >= 75:
                ws_split.cell(row=best_row, column=remark_col).value = "O"
                ws_split.cell(row=best_row, column=best_col).fill = orange_fill
                matched_rows.add(best_row)
                log_info(f'匹配成功: [{mod_title}] (得分:{int(best_score)}%) -> 第 {best_row} 行')
            else:
                log_error(f'匹配失败: [{mod_title}] (最高得分:{int(best_score)}%)')

        # 3. 如果某个三级模块被标记了 O，且该行属于某个大的分组，考虑是否标记同组？
        # 用户反馈“增加了”，可能是因为之前开启了层级传播。现在我们只标最优匹配。

        final_count = len(matched_rows)
        log_info(f'📊 标注完成，共标记 {final_count} 行。')
        log_info('=' * 50 + '\n')

        if progress_callback:
            progress_callback(60, f"已在 P 列标注 {final_count} 个优化模块行")
        return final_count

    @staticmethod
    def fill_manday(ws_calc, manday, progress_callback=None, log_callback=None):
        """
        在结果计算工作表的A22填写送审人天
        """
        if log_callback:
            log_callback(f"正在填写送审人天: {manday}")

        try:
            ws_calc.cell(row=22, column=1).value = manday
            if progress_callback:
                progress_callback(70, f"已填写送审人天: {manday}")
            if log_callback:
                log_callback(f"已在A22填写送审人天: {manday}")
        except Exception as e:
            if log_callback:
                log_callback(f"填写送审人天失败: {str(e)}")
            raise

    @staticmethod
    def process_evaluation(
        architecture_doc_path,
        evaluation_excel_path,
        architecture_text=None,
        manday=0,
        is_manual_mode=False,
        is_auto_evaluate=False,
        project_name=None,
        progress_callback=None,
        log_callback=None,
        verification_callback=None,
    ):
        """
        执行完整的评估流程
        返回: 输出文件路径
        """
        try:
            if progress_callback:
                progress_callback(5, "开始处理...")

            # 1. 解析优化模块
            optimization_modules = []
            if architecture_text:
                optimization_modules = EvaluationProcessor.parse_manual_input(
                    architecture_text, progress_callback, log_callback
                )
            elif architecture_doc_path and not architecture_doc_path.lower().endswith(('.xlsx', '.xlsm')):
                # Word 提取兼容
                optimization_modules = EvaluationProcessor.parse_manual_input(
                    architecture_doc_path, progress_callback, log_callback
                )

            # 2. 加载评估报告Excel
            if log_callback:
                log_callback("正在加载评估报告...")

            from openpyxl import load_workbook
            wb = load_workbook(evaluation_excel_path)

            if progress_callback:
                progress_callback(20, "已加载评估报告")

            # 3. 查找"功能点拆分表"工作表
            ws_split = EvaluationProcessor.find_sheet_by_names(
                wb, ["功能点拆分表", "2、功能点拆分表", "拆分表"]
            )

            if not ws_split:
                raise ValueError("未找到'功能点拆分表'工作表")

            # --- 数据迁移 ---
            if architecture_doc_path and architecture_doc_path.lower().endswith(('.xlsx', '.xlsm')):
                if log_callback: log_callback("检测到拆分表源文件，正在迁移 A-O 列数据...")
                success = EvaluationProcessor._migrate_split_table_data(architecture_doc_path, ws_split, log_callback)
                if success and log_callback: log_callback("✅ A-O 列数据迁移搬运完成")
            if log_callback:
                log_callback(f"找到工作表: {ws_split.title}")

            # 4. 匹配并标注优化模块
            if optimization_modules:
                EvaluationProcessor.match_and_mark_modules(
                    ws_split, optimization_modules, progress_callback=progress_callback, log_callback=log_callback
                )

            # 5. 自动评估规则应用
            if is_auto_evaluate:
                EvaluationProcessor.auto_evaluate_cosmic(
                    ws_split,
                    progress_callback=progress_callback,
                    log_callback=log_callback,
                    verification_callback=verification_callback
                )

            # 6. 查找"结果计算"工作表并填写送审人天
            ws_calc = EvaluationProcessor.find_sheet_by_names(
                wb, ["结果计算", "4、结果计算", "计算"]
            )

            if ws_calc:
                EvaluationProcessor.fill_manday(
                    ws_calc, manday, progress_callback, log_callback
                )

            # 7. 保存文件
            if progress_callback:
                progress_callback(85, "正在保存文件...")

            # 设置日志项目的名称，以便生成带名称的日志文件
            if project_name:
                from utils.runtime_logger import RuntimeLogger
                RuntimeLogger.set_project(project_name)

            raw_base_name = os.path.splitext(os.path.basename(evaluation_excel_path))[0]

            # --- 智能名称清理逻辑 ---
            import re

            # 提取架构文件名作为优先备选 (用户反馈强烈要求使用架构文档名)
            arc_name = ""
            if architecture_doc_path:
                arc_name = os.path.splitext(os.path.basename(architecture_doc_path))[0]

            # 1. 优先使用架构文件的名字 (即“图中红框圈起来的名字”)
            if arc_name and "xxxx" not in arc_name.lower() and len(arc_name) > 2:
                clean_name = arc_name
            # 2. 其次使用传入的项目名称
            elif project_name and "xxxx" not in project_name.lower() and len(project_name) > 2:
                clean_name = project_name
            else:
                # 3. 如果都没有，使用拆分表文件的名称清理
                clean_name = re.sub(r"^附件\s*\d+\s*", "", raw_base_name)
                # 去掉后缀 "拆分表" 及其后面的文字
                clean_name = re.sub(r"功能点拆分表.*$", "", clean_name)
                clean_name = re.sub(r"拆分表.*$", "", clean_name)
                # 去掉前后的空格和特殊符号
                clean_name = clean_name.strip(" -_")

            if not clean_name: clean_name = raw_base_name

            output_name = f"{clean_name}-评估报告.xlsx"

            if architecture_doc_path and os.path.exists(architecture_doc_path):
                output_dir = os.path.dirname(os.path.abspath(architecture_doc_path))
            else:
                output_dir = os.path.dirname(os.path.abspath(evaluation_excel_path))

            output_path = os.path.join(output_dir, output_name)

            # 使用安全保存逻辑
            actual_saved_path = EvaluationProcessor.safe_save(wb, output_path, log_callback)

            if progress_callback:
                progress_callback(100, "处理完成")
            if log_callback:
                log_callback(f"文件已处理完毕")

            return actual_saved_path

        except Exception as e:
            if log_callback:
                log_callback(f"处理失败: {str(e)}")
            raise

    @staticmethod
    def _migrate_split_table_data(src_path, target_ws, log_callback=None):
        """搬运原始拆分表 A-O 列数据 (包含合并单元格、下拉框、公式处理)"""
        try:
            import openpyxl
            from copy import copy
            # 不使用 data_only=True 以便保留公式和下拉框逻辑
            src_wb = openpyxl.load_workbook(src_path, data_only=False)
            src_ws = EvaluationProcessor.find_sheet_by_names(src_wb, ["2、功能点拆分表", "功能点拆分表", "拆分表"])
            if not src_ws: return False

            # 1. 复制合并单元格状态
            # 0. 先清理目标表 A-O 列的合并，避免表头冲突
            try:
                target_merges = list(target_ws.merged_cells.ranges)
                for m_range in target_merges:
                    if m_range.min_col <= 15:
                        target_ws.unmerge_cells(str(m_range))
            except: pass

            for merged_range in src_ws.merged_cells.ranges:
                if merged_range.min_col <= 15:
                    m_min_col = merged_range.min_col
                    m_max_col = min(merged_range.max_col, 15)
                    m_min_row = merged_range.min_row
                    m_max_row = merged_range.max_row
                    try:
                        target_ws.merge_cells(
                            start_row=m_min_row,
                            start_column=m_min_col,
                            end_row=m_max_row,
                            end_column=m_max_col
                        )
                    except: pass

            # 2. 复制数据有效性 (下拉框)
            try:
                for dv in src_ws.data_validations.dataValidation:
                    new_dv = copy(dv)
                    target_ws.add_data_validation(new_dv)
            except: pass

            # 3. 搬运数值、公式与样式
            # 探测实际行数
            last_valid_row = 5
            empty_row_count = 0
            scan_limit = 10000

            # 存储原始 复用度 和 CFP 供评估引擎参考
            original_l_m_values = {}

            for r in range(1, scan_limit): # 全量同步，包含 1-4 行表头
                row_has_data = False

                # 复制行高
                try:
                    if r in src_ws.row_dimensions:
                        h = src_ws.row_dimensions[r].height
                        if h: target_ws.row_dimensions[r].height = h
                except: pass

                for c in range(1, 16): # 搬运范围扩大到 O 列 (1-15)
                    src_cell = src_ws.cell(row=r, column=c)
                    target_cell = target_ws.cell(row=r, column=c)

                    val = src_cell.value
                    if val is not None:
                        target_cell.value = val
                        row_has_data = True
                        if c == 12 or c == 13 or c == 14: # 记录原始 L/M/N 值
                            original_l_m_values[(r, c)] = val

                    if src_cell.has_style:
                        try:
                            # 复制基础样式
                            target_cell.font = copy(src_cell.font)
                            target_cell.border = copy(src_cell.border)
                            target_cell.fill = copy(src_cell.fill)
                            target_cell.alignment = copy(src_cell.alignment)
                            target_cell.number_format = copy(src_cell.number_format)
                        except: pass

                if row_has_data:
                    last_valid_row = r
                    empty_row_count = 0
                else:
                    empty_row_count += 1

                if empty_row_count > 100 and r > src_ws.max_row:
                    break

            # 挂载原始值
            target_ws._original_l_m = original_l_m_values

            if log_callback: log_callback(f"✅ 成功搬运 A-O 列数据及下拉框，共 {last_valid_row} 行。")
            return True
        except Exception as e:
            if log_callback: log_callback(f"Migrate error: {str(e)}")
            return False

    # @staticmethod
    # def auto_evaluate_cosmic(ws_split, progress_callback=None, log_callback=None, verification_callback=None):
    #     """
    #     自动化评估引擎 - 逻辑修正版
    #     1. 过程分类：强分类 EW 和 ERX。
    #     2. 移动映射：根据 H 列动作关键字智能修正 I 列类型。
    #     3. 复用逻辑：核心遵循“ERX 只对 ERX，EW 只对 EW”。
    #     """
    #     # 初始化AI模型预测器
    #     predictor = None
    #     try:
    #         # 加载配置
    #         from extend.matcher_config import MatcherConfig
    #         config_obj = MatcherConfig.load()
    #         use_ai_model = config_obj.get("automation", {}).get("enable_ai_model", True)
    #
    #         if use_ai_model:
    #             from utils.model_predictor import get_predictor
    #             predictor = get_predictor()
    #             # ============ 🔧 模型诊断代码（插入位置：predictor = get_predictor() 之后） ============
    #             if log_callback:
    #                 log_callback("\n" + "=" * 60)
    #                 log_callback("🔧 [模型诊断] 检查 AI 预测器状态...")
    #
    #             if predictor is None:
    #                 if log_callback:
    #                     log_callback("❌ predictor 为 None，模型未初始化")
    #             else:
    #                 if log_callback:
    #                     log_callback(f"✅ predictor 对象: {type(predictor).__name__}")
    #
    #                 # 检查三个核心分类器
    #                 checks = [
    #                     ("legacy_classifier", "利旧/复用/新增分类器"),
    #                     ("move_classifier", "数据移动类型分类器"),
    #                     ("pattern_classifier", "过程模式分类器")  # 🔥 关键检查项
    #                 ]
    #
    #                 for attr_name, desc in checks:
    #                     attr_val = getattr(predictor, attr_name, None)
    #                     status = "✅" if attr_val else "❌"
    #                     type_info = type(attr_val).__name__ if attr_val else "None"
    #                     if log_callback:
    #                         log_callback(f"{status} {desc} [{attr_name}]: {type_info}")
    #
    #                 if log_callback:
    #                     log_callback("=" * 60 + "\n")
    #             # ============ 🔧 诊断代码结束 ============
    #             # if log_callback:
    #             #     log_callback("✅ AI模型加载成功，将使用模型进行智能评估")
    #     except Exception as e:
    #         if log_callback:
    #             log_callback(f"⚠️ AI模型加载失败: {str(e)}，将使用关键词匹配")
    #         predictor = None
    #
    #     if log_callback:
    #         log_callback("🚀 启动自动化评估引擎 (规则强化版)...")
    #
    #     l1_col, l2_col, l3_col, proc_col, desc_col, move_col, remark_col = 2, 3, 4, 7, 8, 9, 16
    #     start_row = 3
    #
    #     # ... (识别逻辑) ...
    #     all_results_for_dialog = [] # 用于验证对话框
    #
    #     # 0. 重新探测最大行数，确保循环完整 (应对某些表格 max_row 不准的问题)
    #     max_r = start_row
    #     # 针对超大文件，扩大探测范围，直接探测到 Excel 的较大切片
    #     effective_max = ws_split.max_row if ws_split.max_row > 10000 else 10000
    #     search_range = min(effective_max + 10000, 1048576) # Excel 最大行数
    #     for i in range(search_range, start_row, -1):
    #         # 每隔 1000 行打印一次探测日志，避免超大文件卡死
    #         if i % 5000 == 0 and log_callback:
    #             # 仅用于后台调试
    #             pass
    #         if ws_split.cell(row=i, column=desc_col).value or ws_split.cell(row=i, column=proc_col).value:
    #             max_r = i
    #             break
    #     if log_callback: log_callback(f"  🔍 评估引擎探测到有效数据行数: {max_r}")
    #
    #     # 0. 加载评估规则关键字 (从配置加载)
    #     rules = config_obj.get("evaluation_rules", {})
    #     ew_keywords = rules.get("ew_group_keywords", ["新增", "删除", "修改", "导入", "上传", "创建", "保存", "更新", "同步", "处理", "注销", "重置", "通过", "拒绝", "配置", "变更"])
    #     erx_keywords = rules.get("erx_group_keywords", ["查看", "查询", "下载", "导出", "搜索", "浏览", "列表", "获取", "详情", "展示", "显示", "统计", "分析", "报表", "图表"])
    #
    #     # 加载分类复用关键字
    #     ew_reuse_kws = rules.get("ew_reuse_keywords", [])
    #     erx_reuse_kws = rules.get("erx_reuse_keywords", [])
    #     reuse_threshold = config_obj.get("reuse_threshold", 70)
    #
    #     entry_kws = rules.get("entry_keywords", ["接收", "送入", "输入", "点击", "发起", "上传", "跳转"])
    #     read_kws = rules.get("read_keywords", ["读取", "检索", "查询", "获取", "通过查询"])
    #     exit_kws = rules.get("exit_keywords", ["返回", "响应", "输出", "导出", "下载", "推送", "回显", "显示", "展示"])
    #     write_kws = rules.get("write_keywords", ["写入", "更新", "保存", "存储", "修改", "逻辑处理", "删除", "新增", "创建", "同步", "注销", "移除", "清除"])
    #
    #     # 1. 分组与类型智能修正
    #     groups = []
    #     current_group = None
    #     current_l3 = None
    #
    #     # 优化模块状态跟踪
    #     curr_l1_o = False
    #     curr_l2_o = False
    #     curr_l3_o = False
    #
    #     for r in range(start_row, max_r + 1):
    #         # 获取层级值
    #         l1_val = ws_split.cell(row=r, column=l1_col).value
    #         l2_val = ws_split.cell(row=r, column=l2_col).value
    #         l3_val = ws_split.cell(row=r, column=l3_col).value
    #         remark_val = str(ws_split.cell(row=r, column=remark_col).value or "").strip()
    #
    #         # 状态继承逻辑：如果该行有模块名且标注了 O，则该层级进入优化状态
    #         if l1_val: curr_l1_o = (remark_val == "O")
    #         if l2_val: curr_l2_o = (remark_val == "O")
    #         if l3_val: curr_l3_o = (remark_val == "O")
    #
    #         # 当前行是否属于优化模块 (继承父级或自身标记)
    #         row_is_o = curr_l1_o or curr_l2_o or curr_l3_o or (remark_val == "O")
    #
    #         if l3_val: current_l3 = str(l3_val).strip()
    #
    #         proc_val = ws_split.cell(row=r, column=proc_col).value
    #         desc_val = str(ws_split.cell(row=r, column=desc_col).value or "").strip()
    #         move_val = str(ws_split.cell(row=r, column=move_col).value or "").upper().strip()
    #
    #         if proc_val and str(proc_val).strip():
    #             current_group = {
    #                 "l3": current_l3,
    #                 "name": str(proc_val).strip(),
    #                 "rows": []
    #             }
    #             groups.append(current_group)
    #         elif desc_val and not current_group:
    #             # 兼容性修复：如果发现了子过程描述但还没开始功能过程分组，
    #             # 自动为这些孤立行建立一个默认分组 (通常取三级模块名)
    #             group_name = current_l3 if current_l3 else f"未定义过程_{r}"
    #             current_group = {
    #                 "l3": current_l3,
    #                 "name": group_name,
    #                 "rows": []
    #             }
    #             groups.append(current_group)
    #
    #         if current_group and desc_val:
    #             # 记录详细匹配信息供后续 infer 使用
    #             # 预提取非功能/利旧信息
    #             legacy_kws = rules.get("legacy_keywords", ["利旧", "内存", "缓存", "校验", "由于"])
    #             reuse_kws = rules.get("reuse_keywords", ["复用", "继承", "沿用"])
    #
    #             detail_reason = ""
    #             from utils.evaluation_processor import COSMICEvaluator
    #             nf_match = COSMICEvaluator.get_non_functional_match(desc_val)
    #             if nf_match:
    #                 detail_reason = f"识别到非功能关键字'{nf_match}'"
    #             else:
    #                 for kw in legacy_kws:
    #                     if kw in desc_val:
    #                         detail_reason = f"识别到关键字'{kw}'"
    #                         break
    #                 if not detail_reason:
    #                     for kw in reuse_kws:
    #                         if kw in desc_val:
    #                             detail_reason = f"识别到复用关键字'{kw}'"
    #                             break
    #
    #             # 动作识别与修正 (优先使用AI模型，关键词作为辅助)
    #             ai_recognize = ""
    #             h_text = desc_val.strip().lower()
    #             import re
    #             h_text = re.sub(r"^[0-9.\-\s]+", "", h_text)
    #
    #             matched_categories = []
    #             matched_kws = []
    #
    #             # --- 策略：AI模型优先，关键词辅助 ---
    #             if predictor:
    #                 # ===== AI模型预测分支 =====
    #                 # 准备上下文信息
    #                 sub_idx = len(current_group["rows"])
    #                 prev_desc = current_group["rows"][-1]["desc"] if sub_idx > 0 else ""
    #                 next_desc = ""  # 需要lookahead，暂时留空
    #
    #                 # 功能过程上下文
    #                 process_context = " ".join([row["desc"] for row in current_group["rows"]])
    #                 process_name = current_group.get("name", "")
    #
    #                 # 步骤1：预测复用类别（利旧/复用/新增）
    #                 reuse_cat, reuse_conf, reuse_reason = predictor.predict_legacy(
    #                     description=desc_val,
    #                     function_name=process_name,
    #                     level3_module=current_l3,
    #                     prev_desc=prev_desc,
    #                     process_context=process_context,
    #                     position_in_group=sub_idx,
    #                     group_size=1  # 暂时使用1，实际需要预先扫描
    #                 )
    #
    #                 if reuse_cat == "LEIJU":
    #                     # 利旧类型
    #                     ai_recognize = "n"
    #                     matched_kws = [f"AI模型判定(置信度{reuse_conf:.2f}): {reuse_reason}"]
    #                     detail_reason = reuse_reason
    #                 else:
    #                     # 预判功能过程类型（基于过程名称关键词）
    #                     process_name_lower = process_name.lower()
    #                     # 优先级1：通知/发送类 → ERX型
    #                     if any(kw in process_name_lower for kw in ["通知", "发送", "推送", "短信", "邮件", "消息"]):
    #                         predicted_proc_type = "ERX"
    #                     elif any(kw in process_name_lower for kw in erx_keywords):
    #                         predicted_proc_type = "ERX"
    #                     elif any(kw in process_name_lower for kw in ew_keywords):
    #                         predicted_proc_type = "EW"
    #                     else:
    #                         predicted_proc_type = ""  # 未知，让模型自行判断
    #
    #                     # 步骤2：预测数据移动类型
    #                     move_type, move_conf, move_reason = predictor.predict_move_type(
    #                         description=desc_val,
    #                         function_name=process_name,
    #                         level3_module=current_l3,
    #                         reuse_cat=reuse_cat,
    #                         prev_desc=prev_desc,
    #                         process_context=process_context,
    #                         position_in_group=sub_idx,
    #                         group_size=1,
    #                         process_type=predicted_proc_type
    #                     )
    #                     ai_recognize = move_type
    #                     matched_categories = [move_type] if move_type != "n" else []
    #                     matched_kws = [f"AI模型判定(置信度{move_conf:.2f}): {move_reason}"]
    #                     detail_reason = f"{reuse_cat} - {move_reason}"
    #             else:
    #                 # ===== 关键词匹配分支（回退/辅助） =====
    #                 for kw in entry_kws:
    #                     if kw in h_text:
    #                         matched_categories.append("E")
    #                         matched_kws.append(kw)
    #                 for kw in read_kws:
    #                     if kw in h_text:
    #                         matched_categories.append("R")
    #                         matched_kws.append(kw)
    #                 for kw in exit_kws:
    #                     if kw in h_text:
    #                         matched_categories.append("X")
    #                         matched_kws.append(kw)
    #                 for kw in write_kws:
    #                     if kw in h_text:
    #                         matched_categories.append("W")
    #                         matched_kws.append(kw)
    #
    #                 # 兜底模糊匹配
    #                 if not matched_categories:
    #                     for kw in ["请求", "提交", "选择", "勾选"]:
    #                         if kw in h_text:
    #                             matched_categories.append("E")
    #                             matched_kws.append(kw)
    #                     for kw in ["提示", "结果"]:
    #                         if kw in h_text:
    #                             matched_categories.append("X")
    #                             matched_kws.append(kw)
    #
    #                 # 去重保留顺序
    #                 matched_categories = list(dict.fromkeys(matched_categories))
    #                 matched_kws = list(dict.fromkeys(matched_kws))
    #
    #                 if matched_categories:
    #                     ai_recognize = matched_categories[0]
    #                 else:
    #                     ai_recognize = "n"
    #
    #             # 获取 Excel 原值
    #             excel_move = move_val if move_val in ["E", "R", "X", "W"] else ""
    #
    #             # 用户要求：思考过程要展示所有可能的类型 E/X
    #             all_ai_types = "/".join(matched_categories) if matched_categories else ai_recognize
    #
    #             # 协调逻辑：如果 Excel 有原值，优先使用Excel值；否则使用AI判断
    #             if excel_move:
    #                 final_move = excel_move
    #             else:
    #                 final_move = ai_recognize
    #
    #             # 标记是否为真正的利旧（只有LEIJU才是真利旧）
    #             is_real_leiju = (ai_recognize == "n" and "LEIJU" in detail_reason) if detail_reason else False
    #
    #             current_group["rows"].append({
    #                 "row": r,
    #                 "desc": desc_val,
    #                 "move": final_move,
    #                 "ai_move": ai_recognize, # 记录 AI 的最终确定动作
    #                 "all_ai_types": all_ai_types, # 记录所有可能的匹配类型 e.g. "E/X"
    #                 "matched_categories": matched_categories, # 原始列表
    #                 "matched_kws": matched_kws, # 记录识别到的关键字
    #                 "excel_move": excel_move, # 记录原始 Excel 值
    #                 "is_optimization": row_is_o,  # 存储继承的状态
    #                 "detail_reason_pre": detail_reason, # 预存识别理由
    #                 "is_real_leiju": is_real_leiju # 标记是否为真正的利旧
    #             })
    #
    #     # 2. 基准池 (改为全局过程池，支持跨模块复用)
    #     global_process_benchmarks = [] # 过程池：[{"name": name, "rows": group_results, "type": proc_type}]
    #
    #     # --- 步骤 1: 独立分析所有过程，确认过程模式类型 (不考虑复用) ---
    #     all_independent_analyses = []
    #
    #     # ============ 🔧 详细诊断：检查过程模式分类器状态 ============
    #     if log_callback:
    #         # log_callback(f"\n{'=' * 60}")
    #         # log_callback(f"[诊断] predictor 状态：{'已加载' if predictor is not None else '未加载'}")
    #         print(f"\n{'=' * 60}")
    #         print(f"[诊断] predictor 状态：{'已加载' if predictor is not None else '未加载'}")
    #
    #         if predictor:
    #             print(f"[诊断] pattern_classifier: {'已加载' if predictor.pattern_classifier else 'None'}")
    #             print(
    #                 f"[诊断] pattern_classifier 类型: {type(predictor.pattern_classifier).__name__ if predictor.pattern_classifier else 'N/A'}")
    #
    #             # pattern_cls_status = "已加载" if predictor.pattern_classifier is not None else "未加载 (None)"
    #             # pattern_cls_type = type(
    #             #     predictor.pattern_classifier).__name__ if predictor.pattern_classifier else "N/A"
    #             # log_callback(f"[诊断] pattern_classifier 状态：{pattern_cls_status}")
    #             # log_callback(f"[诊断] pattern_classifier 类型：{pattern_cls_type}")
    #
    #             if predictor.pattern_classifier:
    #                 log_callback(f"  ✅ 过程模式分类器已加载，将使用 AI 模型预测")
    #             else:
    #                 log_callback(f"  ⚠️ 过程模式分类器为 None，将使用关键词回退")
    #         else:
    #             log_callback(f"  ⚠️ 预测器为 None，将使用关键词回退")
    #
    #         # log_callback(f"[诊断] groups 数量：{len(groups)}")
    #         # log_callback(f"{'=' * 60}\n")
    #         print(f"[诊断] groups 数量: {len(groups)}")
    #         print(f"{'=' * 60}\n")
    #     # ============ 🔧 诊断代码结束 ============
    #
    #     # # 诊断日志：检查过程模式分类器状态
    #     # if log_callback:
    #     #     if predictor:
    #     #         if predictor.pattern_classifier:
    #     #             log_callback(f"  ✅ 过程模式分类器已加载，将使用AI模型预测")
    #     #         else:
    #     #             log_callback(f"  ⚠️ 过程模式分类器未加载，将使用关键词回退")
    #     #     else:
    #     #         log_callback(f"  ⚠️ 预测器未初始化，将使用关键词回退")
    #
    #     for group in groups:
    #         # ============ 🔧 强制诊断：每个过程的预测调用 ============
    #         print(f"\n[诊断] 处理过程: {group['name']}")
    #         print(f"[诊断] predictor 状态: {predictor is not None}")
    #         # if log_callback:
    #         #     log_callback(f"[诊断] 进入 AI 预测判断：predictor={predictor is not None}, pattern_classifier={predictor.pattern_classifier is not None if predictor else 'N/A'}")
    #
    #         if predictor:
    #             print(f"[诊断] pattern_classifier 状态: {predictor.pattern_classifier is not None}")
    #
    #             # log_callback(
    #             #     f"[诊断] pattern_classifier 状态：{'已加载' if predictor.pattern_classifier is not None else '未加载'}")
    #             # log_callback(f"[诊断] 子过程数量：{len(group['rows'])}")
    #             # if group['rows']:
    #             #     log_callback(f"[诊断] level3: {group['rows'][0].get('level3', '')}")
    #         # ============ 🔧 诊断代码结束 ============
    #
    #         name_lower = group["name"].lower()
    #
    #         # 尝试使用AI模型预测过程模式（优先）
    #         proc_type = ""
    #         essential_moves = []
    #
    #         if log_callback:
    #             log_callback(
    #                 f"[诊断] 进入 AI 预测判断：predictor={predictor is not None}, pattern_classifier={predictor.pattern_classifier is not None if predictor else 'N/A'}")
    #
    #         if predictor and predictor.pattern_classifier:
    #             print(f"  ✅ 进入 AI 预测分支")
    #             # if log_callback:
    #             #     log_callback(f"  ✅ 进入 AI 模型预测分支")
    #             try:
    #                 # 提取所有子过程描述
    #                 sub_descs = [r["desc"] for r in group["rows"]]
    #                 level3 = group["rows"][0].get("level3", "") if group["rows"] else ""
    #
    #                 # # AI模型预测
    #                 # if log_callback:
    #                 #     log_callback(f"  🔍 正在对过程'{group['name']}'进行AI模式预测...")
    #                 #     log_callback(f"  📋 子过程描述：{sub_descs[:2]}...")  # 只显示前 2 个
    #                 print(f"  🔍 调用 predict_process_pattern...")
    #                 pattern, conf, reason = predictor.predict_process_pattern(
    #                     process_name=group["name"],
    #                     sub_process_descriptions=sub_descs,
    #                     level3_module=level3
    #                 )
    #                 print(f"  📊 预测结果: {pattern} (置信度={conf:.2f})")
    #                 print(f"  📋 原因: {reason}")
    #                 # if log_callback:
    #                 #     log_callback(f"  📊 预测结果: {pattern} (置信度={conf:.2f})")
    #                 #     log_callback(f"  📋 原因：{reason}")
    #
    #                 if pattern and conf > 0.5:  # 置信度阈值
    #                     proc_type = pattern
    #                     # 根据模式设置必备动作
    #                     # if pattern == "ERX":
    #                     #     essential_moves = ["E", "R", "X"]
    #                     # elif pattern == "EW":
    #                     #     essential_moves = ["E", "W"]
    #                     # elif pattern == "EXEW":
    #                     #     essential_moves = ["E", "X", "W"]  # 2个E、1个X、1个W
    #                     # elif pattern == "EX":
    #                     #     essential_moves = ["E", "X"]
    #                     # elif pattern == "EXEX":
    #                     #     essential_moves = ["E", "X"]  # 2个E、2个X
    #                     essential_moves = {"ERX": ["E", "R", "X"], "EW": ["E", "W"]}.get(pattern, ["E", "R", "X"])
    #                     print(f"  ✅ AI 预测成功: {group['name']} → {pattern}")
    #                     # if log_callback:
    #                     #     log_callback(f"  ✅ AI预测过程模式: {group['name']} → {pattern} (置信度={conf:.2f})")
    #                 else:
    #                     print(f"  ⚠️ 置信度过低 ({conf:.2f})，回退关键词")
    #                     # if log_callback:
    #                     #     log_callback(f"  ⚠️ 置信度过低({conf:.2f}<0.5)或无结果，使用关键词回退")
    #             except Exception as e:
    #                 print(f"  ❌ 预测异常: {str(e)}")
    #                 # if log_callback:
    #                 #     log_callback(f"  ❌ 过程模式AI预测失败: {str(e)}，使用关键词回退")
    #                 import traceback
    #                 traceback.print_exc()
    #                 # if log_callback:
    #                 #     log_callback(f"  详细错误: {traceback.format_exc()}")
    #         else:
    #             print(
    #                 f"  ⚠️ 未进入 AI 分支: predictor={predictor is not None}, pattern_cls={predictor.pattern_classifier is not None if predictor else 'N/A'}")
    #             # ============ 🔧 诊断代码结束 ============
    #             # if log_callback:
    #             #     log_callback(f"  ⚠️ 未进入 AI 预测分支，直接使用关键词回退")
    #             #     log_callback(
    #             #         f"  📋 原因：predictor={predictor}, pattern_classifier={predictor.pattern_classifier if predictor else 'N/A'}")
    #
    #         # 关键词回退逻辑（当AI模型未加载或预测失败时）
    #         if not proc_type:
    #             if log_callback:
    #                 log_callback(f"  🔄 使用关键词回退判定过程'{group['name']}'的类型...")
    #             # 核心分类逻辑：使用关键词判断
    #             # 优先级1：通知/发送/推送类过程 → ERX型（输出类）- 最高优先级
    #             is_notification = any(kw in name_lower for kw in ["通知", "发送", "推送", "短信", "邮件", "消息"])
    #             is_ew_kw = any(kw in name_lower for kw in ew_keywords)
    #             is_erx_kw = any(kw in name_lower for kw in erx_keywords)
    #
    #             if is_notification:
    #                 # 通知类过程：接收触发-读取配置-发送通知 (E-R-X模式)
    #                 # 即使包含"新增"等EW关键词，也优先判定为ERX
    #                 proc_type, essential_moves = "ERX", ["E", "R", "X"]
    #             elif is_erx_kw and not is_ew_kw:
    #                 # 纯ERX关键词
    #                 proc_type, essential_moves = "ERX", ["E", "R", "X"]
    #             elif is_ew_kw and not is_erx_kw:
    #                 # 纯EW关键词
    #                 proc_type, essential_moves = "EW", ["E", "W"]
    #             elif is_erx_kw and is_ew_kw:
    #                 # 同时包含ERX和EW关键词（但不是通知类），根据实际动作判断
    #                 moves_in_group = "".join([r["move"] for r in group["rows"]])
    #                 if "W" in moves_in_group and "X" not in moves_in_group:
    #                     proc_type, essential_moves = "EW", ["E", "W"]
    #                 else:
    #                     proc_type, essential_moves = "ERX", ["E", "R", "X"]
    #             else:
    #                 # 没有明确关键词，根据实际动作判断
    #                 moves_in_group = "".join([r["move"] for r in group["rows"]])
    #                 if "W" in moves_in_group and "X" not in moves_in_group:
    #                     proc_type, essential_moves = "EW", ["E", "W"]
    #                 else:
    #                     proc_type, essential_moves = "ERX", ["E", "R", "X"]
    #
    #         # 独立推理 (获取最终确定的类型和结果)
    #         # group_results = EvaluationProcessor.infer_group_results(group, proc_type, essential_moves)
    #         # 判断 proc_type 是否来自 AI 预测
    #         proc_type_source = "ai" if (predictor and predictor.pattern_classifier and proc_type) else "rule"
    #         group_results = EvaluationProcessor.infer_group_results(
    #             group, proc_type, essential_moves, proc_type_source=proc_type_source
    #         )
    #         actual_type = group_results[0].get("final_proc_type", proc_type) if group_results else proc_type
    #         all_independent_analyses.append({
    #             "group": group,
    #             "results": group_results,
    #             "type": actual_type
    #         })
    #
    #     # --- 步骤 2: 在确认类型后，统一进行复用判定 ---
    #     global_process_benchmarks = [] # 过程池：[{"name": name, "rows": group_results, "type": actual_type}]
    #     all_results_for_dialog = []
    #
    #     for analysis in all_independent_analyses:
    #         group = analysis["group"]
    #         group_results = analysis["results"]
    #         actual_type = analysis["type"]
    #
    #         print(f"[评估推理] 过程: {group['name']} -> 判定类型: {actual_type}")
    #
    #         # 基于分析后的类型进行 Block 复用判定
    #         best_group_match = None
    #         max_group_sim = 0
    #
    #         for prev_group in global_process_benchmarks:
    #             # 兼容性判断: EW 只能复用 EW; ERX 和 EX 可以互相复用
    #             type_compatible = False
    #             prev_type = prev_group["type"]
    #             if actual_type == "EW" and prev_type == "EW":
    #                 type_compatible = True
    #             elif actual_type in ["ERX", "EX"] and prev_type in ["ERX", "EX"]:
    #                 type_compatible = True
    #
    #             if type_compatible:
    #                 sim = EvaluationProcessor.calculate_similarity(group["name"], prev_group["name"])
    #                 if sim > max_group_sim and sim >= reuse_threshold:
    #                     max_group_sim = sim
    #                     best_group_match = prev_group
    #
    #         if best_group_match:
    #             print(f"  ✨ [全过程复用] {group['name']} -> 匹配到: {best_group_match['name']} ({int(max_group_sim)}%)")
    #
    #             new_group_results = []
    #             for orig_res in group_results:
    #                 r, m = orig_res["row"], orig_res["move"]
    #                 is_o = orig_res.get("is_o", False)
    #                 desc = orig_res.get("desc", "")
    #
    #                 if orig_res["value"] == "n":
    #                     # 如果独立分析已经判定为利旧，保留利旧理由
    #                     val, reason = "n", orig_res["reason"]
    #                 else:
    #                     # 核心改进：按动作类型(Move)匹配，而非物理索引
    #                     move_match = None
    #                     for ref_item in best_group_match["rows"]:
    #                         # 动作类型一致，且参考行是非利旧行(复用、优化或新增)
    #                         if ref_item["move"] == m and ref_item["value"] != "n":
    #                             move_match = ref_item
    #                             break
    #
    #                     if move_match:
    #                         ref_val = move_match["value"]
    #                         target_ref = move_match["row"] if (ref_val == "O" or ref_val is None) else ref_val
    #                         val = "O" if is_o else int(target_ref)
    #                         reason = f"{'优化(O)' if is_o else '复用'}({target_ref}) <- [{best_group_match['name']}]"
    #                     else:
    #                         # 没有找到对应动作的复用行，保持独立分析结果
    #                         val, reason = orig_res["value"], orig_res["reason"]
    #
    #                 new_group_results.append({
    #                     "row": r,
    #                     "value": val,
    #                     "reason": reason,
    #                     "move": m,
    #                     "desc": desc,
    #                     "is_highlight": orig_res.get("is_highlight", False),
    #                     "thinking": f"          思考过程：[复用模式] 系统匹配到参考过程 {best_group_match['name']}，本行处理逻辑：{reason}"
    #                 })
    #             group_results = new_group_results
    #
    #         # 记录到基准池 (始终记录最终确定的类型)
    #         global_process_benchmarks.append({"name": group["name"], "rows": group_results, "type": actual_type})
    #
    #         # --- 收集并打印推理日志 ---
    #         for res in group_results:
    #             res["group_name"] = group["name"]
    #
    #             # 获取该行原有的 Excel 值
    #             r_num = res["row"]
    #             orig_move = str(ws_split.cell(row=r_num, column=move_col).value or "").strip().upper()
    #             orig_l = str(ws_split.cell(row=r_num, column=12).value or "").strip()
    #             res["orig_move"] = orig_move
    #             res["orig_l"] = orig_l
    #
    #             all_results_for_dialog.append(res)
    #             print(f"  行 {res['row']}: {res.get('desc', '')} (动作:{res['move']}) -> 预判: {res['reason']}")
    #             if res.get("thinking"):
    #                 for t_line in res["thinking"].split("\n"):
    #                     print(t_line)
    #             last_summary = res.get("summary_thinking", "")
    #
    #         if last_summary:
    #             print(f"  {last_summary}")
    #
    #     # --- 人工验证环节 ---
    #     if verification_callback and all_results_for_dialog:
    #         if log_callback: log_callback("正在等待人工验证评估结果...")
    #         # 注意：all_results_for_dialog 里的 row_type 需要从 value 转换过来
    #         for res in all_results_for_dialog:
    #             v = res.get("value")
    #             if v == "n": res["row_type"] = "利旧"
    #             elif v == "O": res["row_type"] = "优化"
    #             elif v is None: res["row_type"] = "新增"
    #             else: res["row_type"] = "复用"
    #
    #         # 调用回调并获取最终结果
    #         verified_results = verification_callback(all_results_for_dialog)
    #         if verified_results is not None:
    #             all_results_for_dialog = verified_results
    #         else:
    #             if log_callback: log_callback("⚠️ 人工验证已取消，按模型识别结果继续或退出。")
    #             # 如果返回 None，表示点击了取消。为了响应用户“不取消评估”的预期，
    #             # 这里我们保持 all_results_for_dialog 原样（用AI识别的结果）。
    #             # 如果用户彻底想停，应该在主界面点停止。
    #             pass
    #
    #     # --- 统一写入 Excel ---
    #     from openpyxl.styles import Alignment, PatternFill
    #     from openpyxl.cell.cell import MergedCell
    #     center_alignment = Alignment(horizontal='center', vertical='center')
    #     yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    #
    #     processed_count = 0
    #     for res in all_results_for_dialog:
    #         r_num = res["row"]
    #         # 这里的 row_type 是人工修改后的
    #         rt = res.get("row_type", "利旧")
    #         move_type = res.get("move", "E")
    #
    #         # 是否需要高亮显示 (AI 判定不一致或手动标记)
    #         is_hl = res.get("is_highlight", False)
    #         orig_move_raw = str(ws_split.cell(row=r_num, column=move_col).value or "").strip().upper()
    #         if not is_hl and move_type != orig_move_raw and orig_move_raw != "":
    #             is_hl = True
    #
    #         # I 列：数据移动类型 (只读逻辑)
    #         i_cell = ws_split.cell(row=r_num, column=move_col)
    #         orig_move = str(i_cell.value or "").strip().upper()
    #
    #         # L 列：复用度 (只读逻辑)
    #         l_cell = ws_split.cell(row=r_num, column=12)
    #         orig_l = str(l_cell.value or "").strip()
    #
    #         # P 列：备注栏 - 写入识别结果 (优先格式)
    #         p_cell = ws_split.cell(row=r_num, column=remark_col)
    #
    #         # 【响应用户平需求】最终报告中不再应用黄色高亮，不修改客户的数据移动类型和复用度
    #         # 这里的逻辑仅写入备注信息作为参考
    #
    #         if not isinstance(p_cell, MergedCell):
    #             # 确定 AI 的建议前缀
    #             prefix = ""
    #             ai_rt_zh = res.get("row_type", "新增")
    #             if ai_rt_zh == "利旧":
    #                 prefix = "n"
    #             elif ai_rt_zh == "优化":
    #                 prefix = "O"
    #             elif ai_rt_zh == "复用":
    #                 v = res.get("value")
    #                 # 如果是复用，尽量保留数字类型
    #                 prefix = v if isinstance(v, (int, float)) else "复用"
    #
    #             # 计算备注内容：只有当 AI 的判定与原始 L/I 列不符时，才写备注作为核对证据
    #             l_mismatch = (prefix != "" and str(prefix) != orig_l)
    #             if orig_l == "利旧" and prefix == "n": l_mismatch = False
    #             if orig_l == "新增" and (not prefix or prefix == ""): l_mismatch = False
    #             m_mismatch = (move_type != orig_move)
    #
    #             final_remark = []
    #             if l_mismatch or ai_rt_zh in ["复用", "利旧", "优化"]:
    #                 if l_mismatch: final_remark.append(prefix)
    #                 elif ai_rt_zh in ["复用", "利旧"]: final_remark.append(prefix)
    #
    #             if m_mismatch:
    #                 final_remark.append(move_type)
    #
    #             if final_remark:
    #                 valid_items = [x for x in final_remark if x is not None and str(x).strip() != ""]
    #                 if len(valid_items) == 1:
    #                     val = valid_items[0]
    #                     # 尝试转换为数字 (解决：生成的备注数字是字符串类型，我希望是数字类型)
    #                     try:
    #                         val_str = str(val).strip()
    #                         if val_str.replace('.', '', 1).isdigit():
    #                             p_cell.value = float(val_str) if '.' in val_str else int(val_str)
    #                         else:
    #                             p_cell.value = val_str
    #                     except:
    #                         p_cell.value = str(val)
    #                 elif valid_items:
    #                     # 多个备注项合并为字符串
    #                     p_cell.value = ", ".join([str(x) for x in valid_items])
    #
    #         # 【重要】不再执行 i_cell.value = move_type 或 l_cell.value = rt 的覆盖操作
    #         # 也不再添加 M 列公式，确保客户原始文件“原汁原味”
    #
    #         processed_count += 1
    #
    #     print(f"\n✅ 自动评估完成，共处理 {len(groups)} 个功能过程，成功标记 {processed_count} 行。")
    #
    #     # --- 最终汇总逻辑 (仅设置 N5:N10 计算公式，不修改其他数据) ---
    #     try:
    #         from openpyxl.cell.cell import MergedCell
    #         # 合并 N5 到 N10 用于显示合计
    #         for merged_range in list(ws_split.merged_cells.ranges):
    #             if merged_range.min_col == 14 and merged_range.max_col == 14:
    #                 if merged_range.min_row >= 5 and merged_range.max_row <= 10:
    #                     ws_split.unmerge_cells(str(merged_range))
    #
    #         ws_split.merge_cells(start_row=5, start_column=14, end_row=10, end_column=14)
    #         n_top_cell = ws_split.cell(row=5, column=14)
    #         n_top_cell.value = "=SUM(M:M)"
    #         n_top_cell.alignment = Alignment(horizontal='center', vertical='center')
    #
    #         if log_callback: log_callback("✅ 报告汇总计算公式已更新 (N5:N10)")
    #     except Exception as e:
    #         if log_callback: log_callback(f"汇总公式设置失败: {str(e)}")
    #
    #     return processed_count

    @staticmethod
    def auto_evaluate_cosmic(ws_split, progress_callback=None, log_callback=None, verification_callback=None):
        """
        自动化评估引擎 - 完整增强版
        1. 过程分类：AI 模型预测 + 关键词回退
        2. 移动映射：多关键词识别 + 用户优先策略
        3. 复用逻辑：ERX 只对 ERX，EW 只对 EW
        4. 日志输出：完整记录每个判断过程
        """

        # ==================== 辅助函数：双输出日志 ====================
        def log_message(msg):
            """同时输出到 log_callback 和 print，确保日志可见"""
            if log_callback:
                try:
                    log_callback(msg)
                except:
                    pass
            log_info(msg)

        # ==================== 初始化 AI 模型预测器 ====================
        predictor = None
        try:
            from extend.matcher_config import MatcherConfig
            config_obj = MatcherConfig.load()
            use_ai_model = config_obj.get("automation", {}).get("enable_ai_model", True)

            if use_ai_model:
                from utils.model_predictor import get_predictor
                predictor = get_predictor()

                # ============ 🔧 模型诊断代码 ============
                log_message("\n" + "=" * 60)
                log_message("🔧 [模型诊断] 检查 AI 预测器状态...")

                if predictor is None:
                    log_message("❌ predictor 为 None，模型未初始化")
                else:
                    log_message(f"✅ predictor 对象：{type(predictor).__name__}")

                    checks = [
                        ("legacy_classifier", "利旧/复用/新增分类器"),
                        ("move_classifier", "数据移动类型分类器"),
                        ("pattern_classifier", "过程模式分类器")
                    ]

                    for attr_name, desc in checks:
                        attr_val = getattr(predictor, attr_name, None)
                        status = "✅" if attr_val else "❌"
                        type_info = type(attr_val).__name__ if attr_val else "None"
                        log_message(f"{status} {desc} [{attr_name}]: {type_info}")

                    log_message("=" * 60 + "\n")
        except Exception as e:
            log_message(f"⚠️ AI 模型加载失败：{str(e)}，将使用关键词匹配")
            predictor = None

        log_message("🚀 启动自动化评估引擎 (完整增强版)...")

        # ==================== 列定义与配置加载 ====================
        # 动态探测表头列位置（兼容不同 Excel 模板的列布局）
        l1_col = EvaluationProcessor.find_column(ws_split, ["一级模块", "一级", "L1"], default_col=2)
        l2_col = EvaluationProcessor.find_column(ws_split, ["二级模块", "二级", "L2"], default_col=3)
        l3_col = EvaluationProcessor.find_column(ws_split, ["三级模块", "三级", "L3"], default_col=4)
        proc_col = EvaluationProcessor.find_column(ws_split, ["功能过程", "功能简述", "功能名称"], default_col=7)
        desc_col = EvaluationProcessor.find_column(ws_split, ["子过程描述", "子过程", "过程描述", "功能点拆分描述"], default_col=8)
        move_col = EvaluationProcessor.find_column(ws_split, ["数据移动类型", "数据移动", "移动类型", "动作"], default_col=9)
        reuse_col = EvaluationProcessor.find_column(ws_split, ["复用度", "开发类型", "类型", "模式"], default_col=12)
        remark_col = EvaluationProcessor.find_column(ws_split, ["备注", "Remark", "说明"], default_col=16)

        log_message(f"  📋 列定位: 一级={l1_col}, 二级={l2_col}, 三级={l3_col}, 功能过程={proc_col}, 描述={desc_col}, 移动={move_col}, 复用度={reuse_col}, 备注={remark_col}")

        # 动态探测数据起始行：找到表头行的下一行
        start_row = 5  # 默认第5行（常见4行表头）
        for r in range(1, 15):
            cell_val = ws_split.cell(row=r, column=proc_col).value
            if cell_val and "功能过程" in str(cell_val):
                start_row = r + 1
                break

        all_results_for_dialog = []

        # 探测最大行数
        max_r = start_row
        effective_max = ws_split.max_row if ws_split.max_row > 10000 else 10000
        search_range = min(effective_max + 10000, 1048576)
        for i in range(search_range, start_row, -1):
            if ws_split.cell(row=i, column=desc_col).value or ws_split.cell(row=i, column=proc_col).value:
                max_r = i
                break

        log_message(f"  🔍 评估引擎探测到有效数据行数：{max_r}")

        # 加载评估规则关键字
        rules = config_obj.get("evaluation_rules", {})
        ew_keywords = rules.get("ew_group_keywords",
                                ["新增", "删除", "修改", "导入", "上传", "创建", "保存", "更新", "同步", "处理", "注销",
                                 "重置", "通过", "拒绝", "配置", "变更"])
        erx_keywords = rules.get("erx_group_keywords",
                                 ["查看", "查询", "下载", "导出", "搜索", "浏览", "列表", "获取", "详情", "展示",
                                  "显示", "统计", "分析", "报表", "图表"])
        reuse_threshold = config_obj.get("reuse_threshold", 70)

        entry_kws = rules.get("entry_keywords", ["接收", "送入", "输入", "点击", "发起", "上传", "跳转"])
        read_kws = rules.get("read_keywords", ["读取", "检索", "查询", "获取", "通过查询"])
        exit_kws = rules.get("exit_keywords", ["返回", "响应", "输出", "导出", "下载", "推送", "回显", "显示", "展示"])
        write_kws = rules.get("write_keywords",
                              ["写入", "更新", "保存", "存储", "修改", "逻辑处理", "删除", "新增", "创建", "同步",
                               "注销", "移除", "清除"])

        # ==================== 步骤 1: 分组与数据移动类型识别 ====================
        groups = []
        current_group = None
        current_l3 = None

        curr_l1_o = False
        curr_l2_o = False
        curr_l3_o = False

        for r in range(start_row, max_r + 1):
            l1_val = ws_split.cell(row=r, column=l1_col).value
            l2_val = ws_split.cell(row=r, column=l2_col).value
            l3_val = ws_split.cell(row=r, column=l3_col).value
            remark_val = str(ws_split.cell(row=r, column=remark_col).value or "").strip()

            if l1_val:
                curr_l1_o = (remark_val == "O")
            if l2_val:
                curr_l2_o = (remark_val == "O")
            if l3_val:
                curr_l3_o = (remark_val == "O")

            row_is_o = curr_l1_o or curr_l2_o or curr_l3_o or (remark_val == "O")

            if l3_val:
                current_l3 = str(l3_val).strip()

            proc_val = ws_split.cell(row=r, column=proc_col).value
            desc_val = str(ws_split.cell(row=r, column=desc_col).value or "").strip()
            move_val = str(ws_split.cell(row=r, column=move_col).value or "").upper().strip()

            if proc_val and str(proc_val).strip():
                current_group = {
                    "l3": current_l3,
                    "name": str(proc_val).strip(),
                    "rows": []
                }
                groups.append(current_group)
            elif desc_val and not current_group:
                group_name = current_l3 if current_l3 else f"未定义过程_{r}"
                current_group = {
                    "l3": current_l3,
                    "name": group_name,
                    "rows": []
                }
                groups.append(current_group)

            if current_group and desc_val:
                legacy_kws = rules.get("legacy_keywords", ["利旧", "内存", "缓存", "校验", "由于"])
                reuse_kws = rules.get("reuse_keywords", ["复用", "继承", "沿用"])

                detail_reason = ""
                from utils.evaluation_processor import COSMICEvaluator
                nf_match = COSMICEvaluator.get_non_functional_match(desc_val)
                if nf_match:
                    detail_reason = f"识别到非功能关键字'{nf_match}'"
                else:
                    for kw in legacy_kws:
                        if kw in desc_val:
                            detail_reason = f"识别到关键字'{kw}'"
                            break
                    if not detail_reason:
                        for kw in reuse_kws:
                            if kw in desc_val:
                                detail_reason = f"识别到复用关键字'{kw}'"
                                break

                # ==================== 动作识别：多关键词 + 用户优先 ====================
                ai_recognize = ""
                h_text = desc_val.strip().lower()
                import re
                h_text = re.sub(r"^[0-9.\-\s]+", "", h_text)

                matched_categories = []
                matched_kws = []

                if predictor:
                    # ===== AI 模型预测分支 =====
                    sub_idx = len(current_group["rows"])
                    prev_desc = current_group["rows"][-1]["desc"] if sub_idx > 0 else ""

                    process_context = "   ".join([row["desc"] for row in current_group["rows"]])
                    process_name = current_group.get("name", "")

                    reuse_cat, reuse_conf, reuse_reason = predictor.predict_legacy(
                        description=desc_val,
                        function_name=process_name,
                        level3_module=current_l3,
                        prev_desc=prev_desc,
                        process_context=process_context,
                        position_in_group=sub_idx,
                        group_size=1
                    )

                    if reuse_cat == "LEIJU":
                        ai_recognize = "n"
                        matched_kws = [f"AI 模型判定 (置信度{reuse_conf:.2f}): {reuse_reason}"]
                        detail_reason = reuse_reason
                    else:
                        process_name_lower = process_name.lower()
                        if any(kw in process_name_lower for kw in ["通知", "发送", "推送", "短信", "邮件", "消息"]):
                            predicted_proc_type = "ERX"
                        elif any(kw in process_name_lower for kw in erx_keywords):
                            predicted_proc_type = "ERX"
                        elif any(kw in process_name_lower for kw in ew_keywords):
                            predicted_proc_type = "EW"
                        else:
                            predicted_proc_type = ""

                        move_type, move_conf, move_reason = predictor.predict_move_type(
                            description=desc_val,
                            function_name=process_name,
                            level3_module=current_l3,
                            reuse_cat=reuse_cat,
                            prev_desc=prev_desc,
                            process_context=process_context,
                            position_in_group=sub_idx,
                            group_size=1,
                            process_type=predicted_proc_type
                        )
                        ai_recognize = move_type
                        matched_categories = [move_type] if move_type != "n" else []
                        matched_kws = [f"AI 模型判定 (置信度{move_conf:.2f}): {move_reason}"]
                        detail_reason = f"{reuse_cat} - {move_reason}"
                else:
                    # ===== 关键词匹配分支（回退）=====
                    for kw in entry_kws:
                        if kw in h_text:
                            if "E" not in matched_categories:
                                matched_categories.append("E")
                                matched_kws.append(kw)

                    for kw in read_kws:
                        if kw in h_text:
                            if "R" not in matched_categories:
                                matched_categories.append("R")
                                matched_kws.append(kw)

                    for kw in exit_kws:
                        if kw in h_text:
                            if "X" not in matched_categories:
                                matched_categories.append("X")
                                matched_kws.append(kw)

                    for kw in write_kws:
                        if kw in h_text:
                            if "W" not in matched_categories:
                                matched_categories.append("W")
                                matched_kws.append(kw)

                    if not matched_categories:
                        for kw in ["请求", "提交", "选择", "勾选"]:
                            if kw in h_text:
                                if "E" not in matched_categories:
                                    matched_categories.append("E")
                                    matched_kws.append(kw)
                        for kw in ["提示", "结果"]:
                            if kw in h_text:
                                if "X" not in matched_categories:
                                    matched_categories.append("X")
                                    matched_kws.append(kw)

                    matched_categories = list(dict.fromkeys(matched_categories))
                    matched_kws = list(dict.fromkeys(matched_kws))

                    if matched_categories:
                        ai_recognize = matched_categories[0]
                    else:
                        ai_recognize = "n"

                # 🔧 关键修改：移除"EW 型特殊处理"，让"返回...结果"正常识别为 X
                # 不再强制将 X 改为 W，让 infer_group_results 自然处理

                excel_move = move_val if move_val in ["E", "R", "X", "W"] else ""

                # 🔧 all_ai_types 显示所有匹配的类型（如 "E/X/W"）
                all_ai_types = "/".join(matched_categories) if matched_categories else ai_recognize

                if excel_move:
                    final_move = excel_move
                else:
                    final_move = ai_recognize

                is_real_leiju = (ai_recognize == "n" and "LEIJU" in detail_reason) if detail_reason else False

                current_group["rows"].append({
                    "row": r,
                    "desc": desc_val,
                    "move": final_move,
                    "ai_move": ai_recognize,
                    "all_ai_types": all_ai_types,
                    "matched_categories": matched_categories,
                    "matched_kws": matched_kws,
                    "excel_move": excel_move,
                    "is_optimization": row_is_o,
                    "detail_reason_pre": detail_reason,
                    "is_real_leiju": is_real_leiju
                })

        # ==================== 步骤 2: 过程模式类型预测 ====================
        global_process_benchmarks = []
        all_independent_analyses = []

        # ============ 🔧 过程模式分类器状态诊断 ============
        log_message("\n" + "=" * 60)
        log_message("[诊断] predictor 状态：" + ("已加载" if predictor is not None else "未加载"))

        if predictor:
            log_message(f"[诊断] pattern_classifier: {'已加载' if predictor.pattern_classifier else 'None'}")
            log_message(
                f"[诊断] pattern_classifier 类型：{type(predictor.pattern_classifier).__name__ if predictor.pattern_classifier else 'N/A'}")

            if predictor.pattern_classifier:
                log_message("  ✅ 过程模式分类器已加载，将使用 AI 模型预测")
            else:
                log_message("  ⚠️ 过程模式分类器为 None，将使用关键词回退")
        else:
            log_message("  ⚠️ 预测器为 None，将使用关键词回退")

        log_message(f"[诊断] groups 数量：{len(groups)}")
        log_message("=" * 60 + "\n")
        # ============ 🔧 诊断代码结束 ============

        for group in groups:
            # ============ 🔧 每个过程的预测调用诊断 ============
            log_message(f"\n[诊断] 处理过程：{group['name']}")
            log_message(f"[诊断] predictor 状态：{predictor is not None}")
            if predictor:
                log_message(f"[诊断] pattern_classifier 状态：{predictor.pattern_classifier is not None}")
                log_message(f"[诊断] 子过程数量：{len(group['rows'])}")
                if group['rows']:
                    log_message(f"[诊断] level3: {group['rows'][0].get('l3', '')}")
            # ============ 🔧 诊断代码结束 ============

            name_lower = group["name"].lower()

            proc_type = ""
            essential_moves = []

            log_message(
                f"[诊断] 进入 AI 预测判断：predictor={predictor is not None}, pattern_classifier={predictor.pattern_classifier is not None if predictor else 'N/A'}")

            if predictor and predictor.pattern_classifier:
                log_message("  ✅ 进入 AI 预测分支")
                try:
                    sub_descs = [r["desc"] for r in group["rows"]]
                    level3 = group["rows"][0].get("l3", "") if group["rows"] else ""

                    log_message(f"  🔍 调用 predict_process_pattern...")
                    pattern, conf, reason = predictor.predict_process_pattern(
                        process_name=group["name"],
                        sub_process_descriptions=sub_descs,
                        level3_module=level3
                    )
                    log_message(f"  📊 预测结果：{pattern} (置信度={conf:.2f})")
                    log_message(f"  📋 原因：{reason}")

                    if pattern and conf > 0.5:
                        proc_type = pattern
                        essential_moves = {"ERX": ["E", "R", "X"], "EW": ["E", "W"]}.get(pattern, ["E", "R", "X"])
                        log_message(f"  ✅ AI 预测成功：{group['name']} → {pattern}")
                    else:
                        log_message(f"  ⚠️ 置信度过低 ({conf:.2f})，回退关键词")
                except Exception as e:
                    log_message(f"  ❌ 预测异常：{str(e)}")
                    import traceback
                    traceback.print_exc()
            else:
                log_message(
                    f"  ⚠️ 未进入 AI 分支：predictor={predictor is not None}, pattern_cls={predictor.pattern_classifier is not None if predictor else 'N/A'}")

            # 关键词回退逻辑
            if not proc_type:
                log_message(f"  🔄 使用关键词回退判定过程'{group['name']}'的类型...")

                is_notification = any(kw in name_lower for kw in ["通知", "发送", "推送", "短信", "邮件", "消息"])
                is_ew_kw = any(kw in name_lower for kw in ew_keywords)
                is_erx_kw = any(kw in name_lower for kw in erx_keywords)

                if is_notification:
                    proc_type, essential_moves = "ERX", ["E", "R", "X"]
                elif is_erx_kw and not is_ew_kw:
                    proc_type, essential_moves = "ERX", ["E", "R", "X"]
                elif is_ew_kw and not is_erx_kw:
                    proc_type, essential_moves = "EW", ["E", "W"]
                elif is_erx_kw and is_ew_kw:
                    # 🔧 关键改进：CRUD 类操作即使包含 X 也判定为 EW
                    is_crud_ew = any(kw in name_lower for kw in ["新增", "删除", "修改", "保存", "编辑", "更新", "注销"])
                    moves_in_group = "  ".join([r["move"] for r in group["rows"]])
                    if is_crud_ew or ("W" in moves_in_group and "X" not in moves_in_group):
                        proc_type, essential_moves = "EW", ["E", "W"]
                    else:
                        proc_type, essential_moves = "ERX", ["E", "R", "X"]
                else:
                    is_crud_ew = any(kw in name_lower for kw in ["新增", "删除", "修改", "保存", "编辑", "更新", "注销"])
                    moves_in_group = "  ".join([r["move"] for r in group["rows"]])
                    if is_crud_ew or ("W" in moves_in_group and "X" not in moves_in_group):
                        proc_type, essential_moves = "EW", ["E", "W"]
                    else:
                        proc_type, essential_moves = "ERX", ["E", "R", "X"]

            # 独立推理
            proc_type_source = "ai" if (predictor and predictor.pattern_classifier and proc_type) else "rule"
            group_results = EvaluationProcessor.infer_group_results(
                group, proc_type, essential_moves, proc_type_source=proc_type_source
            )

            actual_type = group_results[0].get("final_proc_type", proc_type) if group_results else proc_type

            all_independent_analyses.append({
                "group": group,
                "results": group_results,
                "type": actual_type
            })

        # ==================== 步骤 3: 复用判定 ====================
        global_process_benchmarks = []
        all_results_for_dialog = []

        for analysis in all_independent_analyses:
            group = analysis["group"]
            group_results = analysis["results"]

            actual_type = analysis.get("type", "ERX")

            log_message(f"[评估推理] 过程：{group['name']} -> 判定类型：{actual_type}")

            best_group_match = None
            max_group_sim = 0

            for prev_group in global_process_benchmarks:
                type_compatible = False
                prev_type = prev_group["type"]
                if actual_type == "EW" and prev_type == "EW":
                    type_compatible = True
                elif actual_type in ["ERX", "EX"] and prev_type in ["ERX", "EX"]:
                    type_compatible = True

                if type_compatible:
                    sim = EvaluationProcessor.calculate_similarity(group["name"], prev_group["name"])
                    if sim > max_group_sim and sim >= reuse_threshold:
                        max_group_sim = sim
                        best_group_match = prev_group

            if best_group_match:
                log_message(
                    f"  ✨ [全过程复用] {group['name']} -> 匹配到：{best_group_match['name']} ({int(max_group_sim)}%)")

                new_group_results = []
                for orig_res in group_results:
                    r, m = orig_res["row"], orig_res["move"]
                    is_o = orig_res.get("is_o", False)
                    desc = orig_res.get("desc", "")

                    if orig_res["value"] == "n":
                        val, reason = "n", orig_res["reason"]
                    else:
                        move_match = next(
                            (ref for ref in best_group_match["rows"] if ref["move"] == m and ref["value"] != "n"), None)
                        if move_match:
                            ref_val = move_match["value"]
                            target_ref = move_match["row"] if (ref_val == "O" or ref_val is None) else ref_val
                            val = "O" if is_o else int(target_ref)
                            reason = f"{'优化 (O)' if is_o else '复用'}({target_ref})   <- [{best_group_match['name']}]"
                        else:
                            val, reason = orig_res["value"], orig_res["reason"]

                    new_group_results.append({
                        "row": r,
                        "value": val,
                        "reason": reason,
                        "move": m,
                        "desc": desc,
                        "is_highlight": orig_res.get("is_highlight", False),
                        "thinking": f"          思考过程：[复用模式] 系统匹配到参考过程 {best_group_match['name']}，本行处理逻辑：{reason}"
                    })
                group_results = new_group_results

            global_process_benchmarks.append({"name": group["name"], "rows": group_results, "type": actual_type})

            # 收集并打印推理日志
            for res in group_results:
                res["group_name"] = group["name"]

                r_num = res["row"]
                orig_move = str(ws_split.cell(row=r_num, column=move_col).value or "").strip().upper()
                orig_l = str(ws_split.cell(row=r_num, column=reuse_col).value or "").strip()
                res["orig_move"] = orig_move
                res["orig_l"] = orig_l

                all_results_for_dialog.append(res)
                log_message(f"  行 {res['row']}: {res.get('desc', '')} (动作:{res['move']}) -> 预判：{res['reason']}")
                if res.get("thinking"):
                    for t_line in res["thinking"].split("\n"):
                        log_message(t_line)
                last_summary = res.get("summary_thinking", "")

                if last_summary:
                    log_message(f"  {last_summary}")

        # ==================== 步骤 4: 人工验证环节 ====================
        if verification_callback and all_results_for_dialog:
            log_message("正在等待人工验证评估结果...")
            for res in all_results_for_dialog:
                v = res.get("value")
                if v == "n":
                    res["row_type"] = "利旧"
                elif v == "O":
                    res["row_type"] = "优化"
                elif v is None:
                    res["row_type"] = "新增"
                else:
                    res["row_type"] = "复用"

            verified_results = verification_callback(all_results_for_dialog)
            if verified_results is not None:
                all_results_for_dialog = verified_results
            else:
                log_message("⚠️ 人工验证已取消，按模型识别结果继续")

        # ==================== 步骤 5: 统一写入 Excel ====================
        from openpyxl.styles import Alignment, PatternFill
        from openpyxl.cell.cell import MergedCell
        center_alignment = Alignment(horizontal='center', vertical='center')
        yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

        processed_count = 0
        for res in all_results_for_dialog:
            r_num = res["row"]
            rt = res.get("row_type", "利旧")
            move_type = res.get("move", "E")

            is_hl = res.get("is_highlight", False)
            orig_move_raw = str(ws_split.cell(row=r_num, column=move_col).value or "").strip().upper()
            if not is_hl and move_type != orig_move_raw and orig_move_raw != "":
                is_hl = True

            i_cell = ws_split.cell(row=r_num, column=move_col)
            orig_move = str(i_cell.value or "").strip().upper()

            l_cell = ws_split.cell(row=r_num, column=reuse_col)
            orig_l = str(l_cell.value or "").strip()

            p_cell = ws_split.cell(row=r_num, column=remark_col)

            if not isinstance(p_cell, MergedCell):
                prefix = ""
                ai_rt_zh = res.get("row_type", "新增")
                if ai_rt_zh == "利旧":
                    prefix = "n"
                elif ai_rt_zh == "优化":
                    prefix = "O"
                elif ai_rt_zh == "复用":
                    v = res.get("value")
                    prefix = v if isinstance(v, (int, float)) else "复用"

                l_mismatch = (prefix != "" and str(prefix) != orig_l)
                if orig_l == "利旧" and prefix == "n":
                    l_mismatch = False
                if orig_l == "新增" and (not prefix or prefix == ""):
                    l_mismatch = False
                m_mismatch = (move_type != orig_move)

                final_remark = []
                if l_mismatch or ai_rt_zh in ["复用", "利旧", "优化"]:
                    if l_mismatch:
                        final_remark.append(prefix)
                    elif ai_rt_zh in ["复用", "利旧"]:
                        final_remark.append(prefix)

                if m_mismatch:
                    final_remark.append(move_type)

                if final_remark:
                    valid_items = [x for x in final_remark if x is not None and str(x).strip() != ""]
                    if len(valid_items) == 1:
                        val = valid_items[0]
                        try:
                            val_str = str(val).strip()
                            if val_str.replace('.', '', 1).isdigit():
                                p_cell.value = float(val_str) if '.' in val_str else int(val_str)
                            else:
                                p_cell.value = val_str
                        except:
                            p_cell.value = str(val)
                    elif valid_items:
                        p_cell.value = ",   ".join([str(x) for x in valid_items])

            processed_count += 1

        log_message(f"\n✅ 自动评估完成，共处理 {len(groups)} 个功能过程，成功标记 {processed_count} 行。")

        # 最终汇总逻辑
        try:
            for merged_range in list(ws_split.merged_cells.ranges):
                if merged_range.min_col == 14 and merged_range.max_col == 14:
                    if merged_range.min_row >= 5 and merged_range.max_row <= 10:
                        ws_split.unmerge_cells(str(merged_range))

            ws_split.merge_cells(start_row=5, start_column=14, end_row=10, end_column=14)
            n_top_cell = ws_split.cell(row=5, column=14)
            n_top_cell.value = "=SUM(M:M)"
            n_top_cell.alignment = center_alignment

            log_message("✅ 报告汇总计算公式已更新 (N5:N10)")
        except Exception as e:
            log_message(f"汇总公式设置失败：{str(e)}")

        return processed_count

class COSMICEvaluator:
    """COSMIC评估规则实现"""

    # 加载动态规则
    try:
        from extend.matcher_config import MatcherConfig
        _config = MatcherConfig.load().get("evaluation_rules", {})
        WRITE_KEYWORDS = _config.get("write_keywords", ["更新", "保存", "删除", "写入", "修改", "新增", "创建", "同步"])
        READ_KEYWORDS = _config.get("read_keywords", ["读取", "检索", "查询", "查看", "获取", "搜索"])
        EXIT_KEYWORDS = _config.get("exit_keywords", ["显示", "展示", "结果", "导出", "下载", "响应"])
        ENTRY_KEYWORDS = _config.get("entry_keywords", ["点击", "输入", "按钮", "发送", "接收", "提交"])
        LEGACY_KEYWORDS = _config.get("legacy_keywords", ["利旧", "内存", "缓存", "校验"])
        NON_FUNCTIONAL_KEYWORDS = _config.get("non_functional", ["数据收集", "清洗", "标注", "模型构建", "微调", "部署", "搬迁", "系统迁移"])
    except:
        # 兜底关键字
        WRITE_KEYWORDS = ["更新", "保存", "删除", "写入", "修改", "新增", "创建", "同步"]
        READ_KEYWORDS = ["读取", "检索", "查询", "查看", "获取", "搜索"]
        EXIT_KEYWORDS = ["显示", "展示", "结果", "导出", "下载", "响应"]
        ENTRY_KEYWORDS = ["点击", "输入", "按钮", "发送", "接收", "提交"]
        LEGACY_KEYWORDS = ["利旧", "内存", "缓存", "校验"]
        NON_FUNCTIONAL_KEYWORDS = ["数据收集", "清洗", "标注", "模型构建", "微调", "部署", "搬迁", "系统迁移"]

    @staticmethod
    def evaluate_function_point(entry_count, exit_count, read_count, write_count):
        """
        根据COSMIC规则计算功能点
        """
        return (entry_count or 0) + (exit_count or 0) + (read_count or 0) + (write_count or 0)

    @staticmethod
    def get_non_functional_match(description):
        """返回匹配到的非功能性关键字"""
        for kw in COSMICEvaluator.NON_FUNCTIONAL_KEYWORDS:
            if kw in description: return kw
        return None

    @staticmethod
    def is_non_functional(description):
        """判断是否为非功能性需求 (利旧)"""
        return COSMICEvaluator.get_non_functional_match(description) is not None

    @staticmethod
    def analyze_row_type(description, context_rows, global_writes=None, global_reads=None):
        """
        分析行类型：新增、复用、利旧
        context_rows: 同一个三级模块下的其他行描述
        global_writes/reads: 整个文档中已发现的操作
        """
        if COSMICEvaluator.is_non_functional(description):
            return "利旧"

        # 校验：如果有校验xx，并且前面有提到查询xx可以算利旧
        if "校验" in description:
            # 简化逻辑：同一模块内有其他操作通常校验算利旧
            if context_rows: return "利旧"

        if any(kw in description for kw in COSMICEvaluator.LEGACY_KEYWORDS):
            return "利旧"

        if "用户界面" in description or "UI" in description.upper():
            return "利旧"

        # 写相关逻辑
        if COSMICEvaluator.is_write_op(description):
            # 内部复用
            for other in context_rows:
                if COSMICEvaluator.is_write_op(other):
                    return "复用"
            # 跨模块复用 (导入等)
            if global_writes and any(kw in description for kw in ["导入", "上传", "同步"]):
                return "复用"
            return "新增"

        # 读相关逻辑
        if COSMICEvaluator.is_read_op(description):
            if "详细查询" in description:
                return "新增"
            # 如果有详细查询且在组内，其他算复用
            has_detailed = any("详细查询" in d for d in context_rows)
            if has_detailed: return "复用"

            # 内部复用
            for other in context_rows:
                if COSMICEvaluator.is_read_op(other):
                    return "复用"

            # 跨模块复用
            if global_reads and any(kw in description for kw in ["查询", "查看", "导出"]):
                return "复用"

            return "新增"

        return "新增"

    @staticmethod
    def is_write_op(description):
        return any(kw in description for kw in COSMICEvaluator.WRITE_KEYWORDS)

    @staticmethod
    def is_read_op(description):
        return any(kw in description for kw in COSMICEvaluator.READ_KEYWORDS)

    @staticmethod
    def get_data_movements(description, row_type):
        """
        根据描述和类型获取 E, X, R, W (严格遵循 ⑤ 方案规则)
        """
        if row_type == "利旧":
            return 0, 0, 0, 0

        e, x, r, w = 0, 0, 0, 0
        is_write = any(kw in description for kw in COSMICEvaluator.WRITE_KEYWORDS)
        is_read = any(kw in description for kw in COSMICEvaluator.READ_KEYWORDS)
        is_external = "接口" in description or "外部" in description or "调用" in description

        if is_write:
            # 写类操作：有erxw只留ew (即 E, W)
            e, w = 1, 1
            if is_external:
                # 外部接口的可以是 exex, exew, 增加 X
                x = 1
        elif is_read:
            # 读类操作：有erxw只留erx
            e, r, x = 1, 1, 1
        else:
            # 默认
            e = 1

        return e, x, r, w

    @staticmethod
    def analyze_function(description):
        """保留原接口兼容性"""
        e, x, r, w = COSMICEvaluator.get_data_movements(description, "新增")
        return e, x, r, w
