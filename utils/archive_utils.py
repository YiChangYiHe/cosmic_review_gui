import os
import zipfile
import tempfile
import shutil
from pathlib import Path
from typing import List
from utils.runtime_logger import log_error, log_info


class ArchiveUtils:
    """压缩包处理工具类"""

    @staticmethod
    def is_archive(file_path: str) -> bool:
        """判断是否为支持的压缩包格式"""
        if not file_path:
            return False
        ext = os.path.splitext(file_path)[1].lower()
        return ext in [".zip", ".rar", ".7z"]

    @staticmethod
    def extract_archive(archive_path: str) -> List[str]:
        """
        递归解压压缩包至临时目录，并返回所有文件的绝对路径。
        只提取 Word 和 Excel 相关文件。
        支持嵌套压缩包（如 zip 包 zip）。
        """
        if not os.path.exists(archive_path):
            return []

        # 检查是否支持
        ext = os.path.splitext(archive_path)[1].lower()
        
        # 扩展支持列表
        if ext not in [".zip", ".rar", ".7z"]:
            return []

        # 创建顶级临时目录
        base_name = os.path.splitext(os.path.basename(archive_path))[0]
        # 用随机后缀防止冲突
        import random
        rand_suffix = random.randint(1000, 9999)
        
        # 🚀 从配置加载解压路径，否则回退到 temp 目录
        try:
            from extend.matcher_config import MatcherConfig
            config = MatcherConfig.load()
            base_temp_dir = config.get("storage", {}).get("extraction")
            if not base_temp_dir:
                base_temp_dir = tempfile.gettempdir()
        except:
            base_temp_dir = tempfile.gettempdir()
            
        root_temp_dir = os.path.join(base_temp_dir, f"cosmic_extract_{base_name}_{rand_suffix}")
        
        if not os.path.exists(root_temp_dir):
            os.makedirs(root_temp_dir)

        # 待处理队列 (路径, 目标目录)
        # 初始只包含用户上传的那个压缩包
        pending_archives = [(archive_path, root_temp_dir)]
        
        extracted_files = []
        processed_archives = set() # 防止循环嵌套或重复处理

        # 广度优先处理
        while pending_archives:
            current_archive, extract_to = pending_archives.pop(0)
            
            if current_archive in processed_archives:
                continue
            processed_archives.add(current_archive)

            curr_ext = os.path.splitext(current_archive)[1].lower()
            
            try:
                # 1. 处理 ZIP
                if curr_ext == ".zip":
                    with zipfile.ZipFile(current_archive, "r") as zip_ref:
                        for member in zip_ref.infolist():
                            try:
                                # 解决中文乱码
                                filename = member.filename.encode("cp437").decode("gbk")
                            except:
                                filename = member.filename
                            
                            safe_name = filename.replace("\\", "/").strip("/")
                            target_path = os.path.join(extract_to, safe_name)
                            
                            if member.is_dir():
                                if not os.path.exists(target_path):
                                    os.makedirs(target_path)
                                continue
                                
                            parent_dir = os.path.dirname(target_path)
                            if not os.path.exists(parent_dir):
                                os.makedirs(parent_dir)
                                
                            with zip_ref.open(member) as source, open(target_path, "wb") as target:
                                shutil.copyfileobj(source, target)
                                
                            # 递归检测
                            if os.path.splitext(target_path)[1].lower() in [".zip", ".rar", ".7z"]:
                                sub_dir = os.path.join(parent_dir, os.path.splitext(os.path.basename(target_path))[0] + "_content")
                                pending_archives.append((target_path, sub_dir))

                # 2. 处理 RAR / 7Z (CLI Fallback)
                elif curr_ext in [".rar", ".7z"]:
                    # 尝试寻找解压工具
                    tool_path = ArchiveUtils._find_extract_tool()
                    if tool_path:
                        import subprocess
                        # 构造命令: "Tool" x "Archive" -o"Output" -y
                        # WinRAR/7-Zip 通用参数: x (extract with full paths), -y (assume yes)
                        # 注意 7-Zip 输出目录参数是 -o{Dir} (无空格)
                        
                        cmd = []
                        if "7z" in tool_path.lower():
                            cmd = [tool_path, "x", current_archive, f"-o{extract_to}", "-y"]
                        else:
                            # 假设是 WinRAR (UnRAR.exe)
                             cmd = [tool_path, "x", current_archive, extract_to + "\\", "-y"]
                        
                        # 运行命令
                        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                        
                        # 扫描解压出来的文件进行递归
                        for r, d, f in os.walk(extract_to):
                            for file in f:
                                full_p = os.path.join(r, file)
                                if full_p == current_archive: continue # 跳过自己
                                if os.path.splitext(full_p)[1].lower() in [".zip", ".rar", ".7z"]:
                                     sub_dir = os.path.join(r, os.path.splitext(file)[0] + "_content")
                                     pending_archives.append((full_p, sub_dir))
                    else:
                        log_info(f'[ARCHIVE] 未找到 WinRAR 或 7-Zip，无法解压: {current_archive}')

            except Exception as e:
                log_error(f'解压失败 {current_archive}: {e}')
        
        # 遍历最终的所有文件
        valid_exts = [".doc", ".docx", ".xls", ".xlsx"]
        for root, dirs, files in os.walk(root_temp_dir):
            for f in files:
                if any(f.lower().endswith(ext) for ext in [".zip", ".rar", ".7z"]):
                    continue
                if any(f.lower().endswith(ext) for ext in valid_exts):
                    extracted_files.append(os.path.abspath(os.path.join(root, f)))

        return extracted_files

    @staticmethod
    def _find_extract_tool():
        """寻找系统中可用的解压工具 (WinRAR 或 7-Zip)"""
        # 1. 检查常用路径
        paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
            r"C:\Program Files\WinRAR\WinRAR.exe", # WinRAR.exe support cmdline too
            r"C:\Program Files\WinRAR\UnRAR.exe",
            r"C:\Program Files (x86)\WinRAR\WinRAR.exe"
        ]
        
        for p in paths:
            if os.path.exists(p):
                return p
        
        # 2. 检查 PATH
        import shutil
        if shutil.which("7z"): return "7z"
        if shutil.which("winrar"): return "winrar"
        if shutil.which("unrar"): return "unrar"
        
        return None

        return extracted_files

    @staticmethod
    def cleanup_temp_dir(temp_dir: str):
        """清理临时目录"""
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
