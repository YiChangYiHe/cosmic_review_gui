#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
结构树磁盘缓存：将第0步构建的层级树 + 全文流持久化为 JSON，
同一 Word 文件（路径+mtime+size 一致）重跑时秒级恢复，跳过完整解析。
缓存文件默认写入项目报告目录（打包后运行时生成，不参与打包）。
"""

import json
import os
from pathlib import Path
from utils.runtime_logger import log_error, log_warn

CACHE_VERSION = 2  # v2: 编号剥离前缀0段，旧缓存编号格式不兼容，强制重建

# item 上由索引构建阶段生成的派生字段：不可 JSON 序列化（set）或可重建，落盘前剔除
VOLATILE_ITEM_KEYS = ("char_set", "_global_idx", "content_list")


def get_cache_path(word_path, project_name=None):
    """缓存 JSON 路径：放在项目根目录下（不带时间戳），确保跨次运行复用"""
    from utils.initial_review_report import project_output_dir
    # 注意：这里只用项目文件夹（不带时间戳），缓存跨次运行复用
    base = project_output_dir(project_name)
    if not base:
        try:
            base = str(Path(word_path).parent)
        except Exception:
            return None
    stem = Path(word_path).stem
    # 缓存文件名固定
    return os.path.join(base, f"{stem}-tree_cache.json")


def _file_meta(word_path):
    try:
        return os.path.getmtime(word_path), os.path.getsize(word_path)
    except OSError:
        return 0, 0


def _strip_volatile(items):
    for it in items:
        if isinstance(it, dict):
            for k in VOLATILE_ITEM_KEYS:
                it.pop(k, None)
    return items


def save_cache(word_path, cached_data, project_name=None):
    """保存缓存。cached_data 需含 items / full_text_content。失败静默。"""
    cache_path = get_cache_path(word_path, project_name)
    if not cache_path:
        return None
    mtime, size = _file_meta(word_path)
    payload = {
        "version": CACHE_VERSION,
        "path": os.path.abspath(word_path),
        "mtime": mtime,
        "size": size,
        "items": _strip_volatile(_jsonable(cached_data.get("items", []))),
        "full_text_content": _jsonable(cached_data.get("full_text_content", [])),
    }
    try:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        tmp = cache_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        os.replace(tmp, cache_path)
        return cache_path
    except Exception as e:
        log_warn(f'  [WARN] 结构树缓存写入失败: {e}')

        return None


def _jsonable(obj):
    """递归转换 set -> list，保证可 JSON 序列化"""
    if isinstance(obj, dict):
        return {k: _jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(v) for v in obj]
    if isinstance(obj, (set, frozenset)):
        return sorted(_jsonable(v) for v in obj)
    if isinstance(obj, (str, int, float, bool)) or obj is None:
        return obj
    return str(obj)


def load_cache(word_path, project_name=None):
    """按 路径+mtime+size+版本 校验缓存，命中返回 payload，未命中返回 None"""
    cache_path = get_cache_path(word_path, project_name)
    if not cache_path or not os.path.exists(cache_path):
        return None
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception as e:
        log_warn(f'  [WARN] 结构树缓存读取失败（将重新解析）: {e}')

        return None
    if not isinstance(payload, dict) or payload.get("version") != CACHE_VERSION:
        return None
    mtime, size = _file_meta(word_path)
    if (
        payload.get("path") != os.path.abspath(word_path)
        or payload.get("mtime") != mtime
        or payload.get("size") != size
    ):
        return None
    return payload
