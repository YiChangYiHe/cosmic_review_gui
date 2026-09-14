import os
import re
import sys
from datetime import datetime


def get_resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    elif getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        internal_path = os.path.join(exe_dir, "_internal", relative_path)
        if os.path.exists(internal_path):
            return internal_path
        base_path = exe_dir
    else:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)


def clean_project_name(filename, strip_version=False):
    if not filename:
        return "未命名项目"

    name = os.path.basename(filename).strip()
    name = re.sub(r"\.(?:docx?|xlsx?|pdf)$", "", name, flags=re.IGNORECASE).strip()

    versions = re.findall(r"([vV]\d+(?:\.\d+)*)", name)
    version_str = versions[-1] if versions else ""

    name = re.sub(r"^附件\s*\d+\s*[：:.\-\s]*\s*", "", name)
    name = re.sub(r"^\d{1,3}[\.．]\s*", "", name)
    name = name.strip()

    truncate_keywords = [
        "产品需求说明书", "需求规格说明书", "需求规格书", "需求说明书",
        "规格说明书", "规格书", "功能点拆分表", "功能拆分表", "拆分表",
        "资产清单", "审计方案", "测试用例", "说明书", "文档", "需求",
    ]
    for kw in truncate_keywords:
        if kw in name:
            parts = name.split(kw)
            prefix = parts[0].strip()
            suffix = parts[1].strip() if len(parts) > 1 else ""
            suffix_clean = re.sub(r"^[\-\_\s:：]+", "", suffix)
            if len(prefix) < 4 and suffix_clean:
                name = suffix_clean
            else:
                name = prefix
            break

    while True:
        prev_name = name
        name = re.sub(r"(?:20)?\d{6,8}\s*$", "", name)
        name = re.sub(r"[\(（]\d+[\)）]\s*$", "", name)
        name = re.sub(r"\s*-\s*副本\s*$", "", name)
        name = re.sub(r"(?:的项目|项目|的需求|的产品|产品|的|之)+$", "", name.strip())
        name = name.strip(" :：-－_|'\"().（）")
        if name == prev_name:
            break

    if version_str and version_str not in name:
        name = f"{name} {version_str}"

    # 【新增】如果要求去除版本号，则在这里处理（用于统一文件夹命名）
    if strip_version:
        name = re.sub(r"\s*[vV]\d+(?:\.\d+)*\s*$", "", name).strip()

    return name.strip() if name.strip() else "未命名项目"


# 【迁移】初评报告路径逻辑已统一迁移到 utils/initial_review_report.py，
# 此处不再提供 prepare_report_paths / get_project_report_dir，
# 初评产物路径一律通过 initial_review_report 模块生成。

def clear_directory(directory_path):
    """彻底清空文件夹下的所有文件和子文件夹"""
    if not os.path.exists(directory_path):
        return True, "文件夹不存在"

    import shutil
    try:
        for filename in os.listdir(directory_path):
            file_path = os.path.join(directory_path, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        return True, "清理成功"
    except Exception as e:
        return False, f"清理失败: {str(e)}"


def open_directory(path):
    """在操作系统中打开文件夹"""
    if not os.path.exists(path):
        return

    import platform
    import subprocess

    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


