import json
import os
import copy
import time
from utils.path_utils import get_resource_path
from utils.runtime_logger import log_error


class MatcherConfig:
    # 配置文件优先使用当前目录，其次用户目录下，确保可写且不会随程序删除而消失
    _LOCAL_CONFIG = os.path.join(os.getcwd(), "matcher_config.json")
    _USER_DIR = os.path.join(os.path.expanduser("~"), ".cosmic_review")
    if not os.path.exists(_USER_DIR):
        os.makedirs(_USER_DIR, exist_ok=True)

    _USER_CONFIG = os.path.join(_USER_DIR, "matcher_config.json")

    # 当前实际使用的配置文件路径
    CONFIG_FILE = _LOCAL_CONFIG if os.path.exists(_LOCAL_CONFIG) else _USER_CONFIG

    # [性能] 短 TTL 缓存：load() 被绘制/动画等高频路径调用（步骤条动画每 50ms
    # 触发多次），每次都读磁盘 + 解析 JSON 会让 UI 线程持续做 IO。
    _CACHE_TTL = 1.0  # 秒
    _load_cache = None  # {"file", "mtime", "ts", "config"}

    @classmethod
    def load(cls):
        # 实时探测文件是否存在，如果本地新创建了则切换到本地
        if os.path.exists(cls._LOCAL_CONFIG):
            cls.CONFIG_FILE = cls._LOCAL_CONFIG
        else:
            cls.CONFIG_FILE = cls._USER_CONFIG

        cfg_file = cls.CONFIG_FILE
        try:
            mtime = os.path.getmtime(cfg_file) if os.path.exists(cfg_file) else None
        except OSError:
            mtime = None

        cached = cls._load_cache
        if (
            cached
            and cached["file"] == cfg_file
            and cached["mtime"] == mtime
            and time.monotonic() - cached["ts"] < cls._CACHE_TTL
        ):
            # 调用方会修改返回值（改完再 save），必须给副本
            return copy.deepcopy(cached["config"])

        config = cls.get_defaults()
        if os.path.exists(cfg_file):
            try:
                with open(cfg_file, "r", encoding="utf-8") as f:
                    user_config = json.load(f)
                    # 深度更新配置
                    cls._deep_update(config, user_config)
            except Exception as e:
                log_error(f'Error loading config: {e}')

        # 统一路径格式
        if "storage" in config:
            for key, path in config["storage"].items():
                if path:
                    config["storage"][key] = os.path.normpath(path)

        cls._load_cache = {
            "file": cfg_file,
            "mtime": mtime,
            "ts": time.monotonic(),
            "config": config,
        }
        return copy.deepcopy(config)

    @staticmethod
    def _deep_update(d, u):
        for k, v in u.items():
            if isinstance(v, dict) and k in d and isinstance(d[k], dict):
                MatcherConfig._deep_update(d[k], v)
            else:
                d[k] = v
        return d

    @classmethod
    def save(cls, config):
        try:
            # 在保存时也检查一下路径
            if os.path.exists(cls._LOCAL_CONFIG):
                cls.CONFIG_FILE = cls._LOCAL_CONFIG

            with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            # 保存后立即失效缓存，让下一次 load 读到最新内容
            cls._load_cache = None
        except Exception as e:
            log_error(f'Error saving config: {e}')

    @classmethod
    def get_defaults(cls):
        # 获取用户文档目录作为默认存放根目录
        user_docs = os.path.normpath(os.path.expanduser("~/Documents/CosmicReview"))
        return {
            "hierarchy": {
                "level1_col": 1,
                "level2_col": 2,
                "level3_col": 3,
                "sheet_name": 2,  # Default to 3rd sheet (index 2)
            },
            "process": {
                "column": 6,  # Default functional process column (7th column, index 6)
                "sheet_name": 2,  # Default to 3rd sheet (index 2)
            },
            "storage": {
                "initial_review": os.path.join(user_docs, "InitialReview"),
                "re_review": os.path.join(user_docs, "ReReview"),
                "receipt": os.path.join(user_docs, "Receipt"),
                "logs": os.path.join(user_docs, "Logs"),
                "extraction": os.path.join(user_docs, "Extraction"),
            },
            "theme": {
                "primary_color": "#2563eb",
                "text_color_light": "#1e293b",
                "text_color_dark": "#f3f4f6",
                "font_family": "Microsoft YaHei UI",
                "font_size": 14,
                "is_dark": False,
            },
            "automation": {
                "auto_open": True,
                "reuse_threshold": 70  # 全篇复用模糊匹配阈值 (0-100)
            },
            "evaluation_folder": "", # 默认评估文件夹
            "evaluation_rules": {
                "write_keywords": ["更新", "保存", "删除", "写入", "修改", "新增", "创建", "同步", "清空", "变更", "录入"],
                "read_keywords": ["读取", "检索", "查询", "查看", "获取", "搜索", "定位", "加载", "统计"],
                "exit_keywords": ["显示", "展示", "结果", "导出", "下载", "响应", "回显", "推送", "返回"],
                "entry_keywords": ["点击", "输入", "按钮", "发送", "接收", "提交", "请求", "触发", "勾选"],
                "ew_group_keywords": ["新增", "删除", "修改", "导入", "上传", "创建", "保存", "更新", "同步", "处理", "注销", "重置", "通过", "拒绝", "配置", "变更"],
                "erx_group_keywords": ["查看", "查询", "下载", "导出", "搜索", "浏览", "列表", "获取", "详情", "展示", "显示", "统计", "分析", "报表", "图表"],
                "ew_reuse_keywords": ["新增", "修改", "删除", "导入"],
                "erx_reuse_keywords": ["查询", "查看", "导出"],
                "non_functional": ["数据收集", "清洗", "标注", "模型构建", "微调", "部署", "搬迁", "系统迁移", "背景", "文案", "像素"],
                "legacy_keywords": ["利旧", "内存", "缓存", "校验", "由于"]
            }
        }
