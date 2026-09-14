import pandas as pd
import os
import re
from datetime import datetime
from extend.matcher_config import MatcherConfig
from utils.runtime_logger import log_error


class ReportGenerator:
    """自动化报表生成工具"""

    @staticmethod
    def _clean_project_name(name):
        """净化项目名，统一调用 path_utils 中的函数"""
        from utils.path_utils import clean_project_name
        return clean_project_name(name)

    @classmethod
    def generate_validation_report(cls, task_name, results, output_dir=None, base_name=None):
        """生成全量比对报表"""
        try:
            from utils.initial_review_report import report_file_path

            details = results.get("all_details", [])
            if not details: return None

            # 【核心修复】使用初评报告统一路径模块，确保进入本次运行的时间戳文件夹
            project_subdir, target_path = report_file_path(
                report_type="模板校验评估报告",
                project_name_or_path=task_name,
                report_base_name=base_name,
            )

            # 如果外部强制传入了特定的 output_dir，则覆盖
            if output_dir:
                target_path = os.path.join(output_dir, "模板校验评估报告.xlsx")
                project_subdir = output_dir

            os.makedirs(project_subdir, exist_ok=True)

            df = pd.DataFrame(details)
            column_map = {"level": "层级", "chapter": "标准章节", "status": "判定结果",
                          "tpl_sentence": "模板参考句 (对照)", "target_sentence": "上传文本句 (对照)",
                          "score": "重合度"}
            existing_cols = [c for c in df.columns if c in column_map]
            df = df[existing_cols].rename(columns=column_map)
            if "重合度" in df.columns:
                df["重合度"] = df["重合度"].apply(lambda x: f"{x * 100:.1f}%" if isinstance(x, (int, float)) else x)

            # 防占用处理
            final_path = target_path
            counter = 1
            while True:
                try:
                    if os.path.exists(final_path):
                        with open(final_path, "a"): pass
                    break
                except (IOError, PermissionError):
                    name, ext = os.path.splitext(os.path.basename(target_path))
                    final_path = os.path.join(project_subdir, f"{name}({counter}){ext}")
                    counter += 1

            # 备份
            if os.path.exists(target_path) and target_path != final_path:
                try:
                    import shutil
                    backup_name = f"模板校验评估报告_bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    shutil.copy2(target_path, os.path.join(project_subdir, backup_name))
                except:
                    pass

            with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="校验详情")
                workbook = writer.book
                worksheet = writer.sheets["校验详情"]
                from openpyxl.styles import Alignment
                for i, col in enumerate(df.columns):
                    column_len = max((df[col].astype(str).map(len).max() if not df[col].empty else 10), len(col)) + 2
                    col_letter = chr(65 + i)
                    worksheet.column_dimensions[col_letter].width = min(column_len, 60)
                    for cell in worksheet[col_letter]:
                        cell.alignment = Alignment(wrap_text=True if "对照" in col else False, vertical="center")

            return final_path
        except Exception as e:
            log_error(f'生成报表失败: {e}')
            return None