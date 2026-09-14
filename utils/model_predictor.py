# utils/model_predictor.py
"""
COSMIC 评估模型预测器 - 匹配训练特征版
功能：
1. 加载训练好的分类模型
2. 提取与训练时一致的特征
3. 使用模型优先、规则辅助的预测策略
"""

import os
import joblib
import re
import threading
import jieba
import jieba.posseg as pseg
import pandas as pd
from pathlib import Path

from utils.runtime_logger import log_error, log_warn
from utils.runtime_logger import log_info


class COSMICModelPredictor:
    """COSMIC 评估模型预测器"""

    def __init__(self, model_dir=None):
        """初始化预测器"""
        if model_dir is None:
            current_dir = Path(__file__).parent.parent
            model_dir = current_dir / "model"

        self.model_dir = Path(model_dir)
        self.legacy_classifier = None
        self.move_classifier = None
        self.pattern_classifier = None  # 过程模式分类器

        self._load_models()

    def _load_models(self):
        """加载训练好的模型"""
        try:
            legacy_model_path = self.model_dir / "cosmic_legacy_classifier.joblib"
            if legacy_model_path.exists():
                self.legacy_classifier = joblib.load(legacy_model_path)
                log_info(f'✅ 成功加载评估模型: {legacy_model_path}')
            else:
                log_warn(f'⚠️ 评估模型不存在: {legacy_model_path}')

            move_model_path = self.model_dir / "cosmic_move_classifier.joblib"
            if move_model_path.exists():
                self.move_classifier = joblib.load(move_model_path)
                log_info(f'✅ 成功加载数据移动模型: {move_model_path}')
            else:
                log_warn(f'⚠️ 数据移动模型不存在: {move_model_path}')

            pattern_model_path = self.model_dir / "cosmic_process_pattern_classifier.joblib"
            if pattern_model_path.exists():
                self.pattern_classifier = joblib.load(pattern_model_path)
                log_info(f'✅ 成功加载过程模式模型: {pattern_model_path}')
            else:
                log_warn(f'⚠️ 过程模式模型不存在: {pattern_model_path}')

        except Exception as e:
            log_error(f'❌ 加载模型失败: {str(e)}')
            raise

    def cosmic_tokenizer(self, text: str) -> str:
        """COSMIC专用分词（与训练保持一致）"""
        pairs = []
        for word, flag in pseg.cut(str(text)):
            word = word.strip()
            if not word:
                continue
            is_v = flag.startswith("v")
            is_n = flag.startswith("n") or flag in ("an", "vn", "nz", "nr", "ns")
            if is_v or is_n:
                pairs.append((word, flag))

        tokens = [w for w, _ in pairs]
        compounds = []
        for i, (word, flag) in enumerate(pairs):
            if flag.startswith("v"):
                # 名词_动词（主谓）
                if i > 0:
                    pw, pf = pairs[i - 1]
                    if pf.startswith("n") or pf in ("nr", "ns", "nz"):
                        compounds.append(f"{pw}_{word}")
                # 动词_名词（动宾）
                for j in range(i + 1, min(i + 3, len(pairs))):
                    nw, nf = pairs[j]
                    if nf.startswith("n") or nf in ("an", "vn", "nz"):
                        compounds.append(f"{word}_{nw}")
                        break
                    if nf.startswith("v") and nf != "vn":
                        break

        all_tokens = tokens + compounds
        return " ".join(all_tokens) if all_tokens else " ".join(jieba.cut(str(text)))

    def extract_features(self, description, function_name="", level3_module="",
                        sub_move_type="", process_move_type="",
                        prev_desc="", next_desc="", process_context="",
                        position_in_group=0, group_size=1, reuse_cat="NEW"):
        """提取完整特征（匹配训练时的特征集）"""
        # 合并文本
        combined_text = f"{description} {function_name}"

        # 分词处理
        tokenized_text = self.cosmic_tokenizer(combined_text)
        tokenized_context = self.cosmic_tokenizer(process_context or description)
        tokenized_process_name = self.cosmic_tokenizer(function_name)
        tokenized_prev = self.cosmic_tokenizer(prev_desc)
        tokenized_next = self.cosmic_tokenizer(next_desc)

        # 位置特征
        position_ratio = position_in_group / max(group_size, 1)
        is_first = 1 if position_in_group == 0 else 0
        is_last = 1 if position_in_group == group_size - 1 else 0
        is_mid = 5.0 if (position_in_group > 0 and position_in_group < group_size - 1) else 0

        # 位置标签
        position_tags = []
        if is_first:
            position_tags.append("__POS_FIRST__")
        elif is_last:
            position_tags.append("__POS_LAST__")
        else:
            position_tags.append("__POS_MID__")

        if group_size == 2:
            position_tags.append("__GS2__")
        elif group_size == 3:
            position_tags.append("__GS3__")
        elif group_size >= 4:
            position_tags.append("__GS4P__")

        position_tags_str = " ".join(position_tags)

        # 利旧信号特征
        leiju_signals = self._extract_leiju_signals(description)

        # 交叉特征
        type_cross = f"{sub_move_type}_IN_{process_move_type}" if sub_move_type and process_move_type else ""

        # 构建特征字典（匹配训练时的列名）
        features = {
            'tokenized_text': tokenized_text,
            'tokenized_context': tokenized_context,
            'tokenized_process_name': tokenized_process_name,
            'tokenized_prev': tokenized_prev,
            'tokenized_next': tokenized_next,
            'level3': level3_module or 'unknown',
            'reuse_cat': reuse_cat,
            'position_ratio': position_ratio,
            'is_first': is_first,
            'is_last': is_last,
            'is_mid': is_mid,
            'is_first_w': is_first * 10.0,
            'is_last_w': is_last * 10.0,
            'group_size': group_size,
            'position_tags': position_tags_str,
            'sub_move_type': sub_move_type or 'E',
            'pmt': process_move_type or 'ERX',
            'leiju_signals': leiju_signals,
            'type_cross': type_cross,
        }

        return features

    def _extract_leiju_signals(self, description):
        """提取利旧信号"""
        text = str(description).lower()

        # 利旧关键词
        leiju_kws = [
            '日志', '审计', '缓存', '内存', 'redis', 'session', 'cookie',
            '界面', '用户界面', '页面', 'ui', 'ue',
            '模糊', '模糊查询', '筛选', '过滤', '去重',
            '计算', '判断', '校验', '验证', '提示', '加密', '解密', '解析', '转化',
            'token', '票据', '令牌', '算法', '模型',
            '配置文件', '国产化', '数据搬迁', '系统迁移',
        ]

        count = sum(1 for kw in leiju_kws if kw in text)
        return str(count)

    def predict_legacy(self, description, function_name="", level3_module="",
                      sub_move_type="", process_move_type="",
                      prev_desc="", next_desc="", process_context="",
                      position_in_group=0, group_size=1):
        """
        预测复用类别 (NEW/REUSE/LEIJU)

        Returns:
            tuple: (reuse_cat: str, confidence: float, reason: str)
        """
        if not self.legacy_classifier:
            return self._rule_based_legacy(description)

        try:
            features = self.extract_features(
                description, function_name, level3_module,
                sub_move_type, process_move_type,
                prev_desc, next_desc, process_context,
                position_in_group, group_size
            )

            # 使用 DataFrame 格式
            feature_df = pd.DataFrame([features])
            prediction = self.legacy_classifier.predict(feature_df)[0]

            if hasattr(self.legacy_classifier, 'predict_proba'):
                proba = self.legacy_classifier.predict_proba(feature_df)[0]
                confidence = max(proba)
            else:
                confidence = 1.0

            reuse_cat = str(prediction)
            reason = f"AI模型判定: {reuse_cat}"

            return reuse_cat, confidence, reason

        except Exception as e:
            log_error(f'模型预测失败，回退到规则判断: {str(e)}')
            import traceback
            traceback.print_exc()
            return self._rule_based_legacy(description)

    def predict_move_type(self, description, function_name="", level3_module="",
                         reuse_cat="NEW", prev_desc="", next_desc="", process_context="",
                         position_in_group=0, group_size=1, process_type=""):
        """
        预测数据移动类型 (E / R / X / W)

        Args:
            process_type: 功能过程类型 (ERX/EW/EXEW等)，用于约束预测结果

        Returns:
            tuple: (move_type: str, confidence: float, reason: str)
        """
        if not self.move_classifier:
            return self._rule_based_move_type(description)

        try:
            features = self.extract_features(
                description, function_name, level3_module,
                "", process_type,  # 传递process_type作为process_move_type
                prev_desc, next_desc, process_context,
                position_in_group, group_size, reuse_cat
            )

            # 使用 DataFrame 格式
            feature_df = pd.DataFrame([features])
            prediction = self.move_classifier.predict(feature_df)[0]

            if hasattr(self.move_classifier, 'predict_proba'):
                proba = self.move_classifier.predict_proba(feature_df)[0]
                confidence = max(proba)
            else:
                confidence = 1.0

            move_type = str(prediction)
            reason = f"AI模型判定: {move_type}"

            return move_type, confidence, reason

        except Exception as e:
            log_error(f'模型预测失败，回退到规则判断: {str(e)}')
            import traceback
            traceback.print_exc()
            return self._rule_based_move_type(description)

    def predict_process_pattern(self, process_name, sub_process_descriptions, level3_module=""):
        """
        预测功能过程模式类型

        Args:
            process_name: 功能过程名称
            sub_process_descriptions: 所有子过程描述列表
            level3_module: 三级模块名

        Returns:
            (pattern, confidence, reason): 模式类型(ERX/EW/EXEW/EX/EXEX), 置信度, 原因
        """
        if self.pattern_classifier is None:
            return "", 0.0, "过程模式分类器未加载"

        try:
            # 聚合所有子过程描述
            process_text = " ".join([str(d) for d in sub_process_descriptions if pd.notna(d)])

            # 分词
            tokenized_process = self.cosmic_tokenizer(process_text)

            # 构建特征
            features = pd.DataFrame([{
                "tokenized_process": tokenized_process,
                "level3_val": level3_module or "unknown"
            }])

            # 预测
            pattern = self.pattern_classifier.predict(features)[0]
            proba = self.pattern_classifier.predict_proba(features)[0]
            confidence = float(proba.max())

            # 生成原因
            reason = f"过程名称'{process_name}'包含{len(sub_process_descriptions)}个子过程"

            return pattern, confidence, reason

        except Exception as e:
            return "", 0.0, f"过程模式预测失败: {str(e)}"

    def _rule_based_legacy(self, description):
        """基于规则的复用类别判断（回退方案）"""
        text = str(description).lower()

        legacy_keywords = ["利旧", "内存", "缓存", "校验", "算法", "计算", "判断",
                          "提示", "加密", "解密", "解析", "转化", "界面", "配置",
                          "日志", "审计", "模糊", "去重", "token"]

        for kw in legacy_keywords:
            if kw in text:
                return "LEIJU", 1.0, f"识别到关键字'{kw}'"

        return "NEW", 1.0, "未匹配到利旧特征"

    def _rule_based_move_type(self, description):
        """基于规则的数据移动类型判断（回退方案）"""
        text = str(description).lower()

        # 写操作
        if any(kw in text for kw in ["新增", "删除", "修改", "保存", "更新", "写入", "存储"]):
            return "W", 1.0, "识别到写操作"

        # 读操作
        if any(kw in text for kw in ["查看", "查询", "详情", "列表", "读取", "检索", "搜索"]):
            return "R", 1.0, "识别到读操作"

        # 返回/输出操作
        if any(kw in text for kw in ["返回", "响应", "输出", "导出", "下载", "展示", "显示"]):
            return "X", 1.0, "识别到输出操作"

        # 输入操作
        if any(kw in text for kw in ["接收", "输入", "点击", "发起", "提交"]):
            return "E", 1.0, "识别到输入操作"

        return "E", 1.0, "默认Entry"

# 全局单例（加锁：并行任务首次同时调用时防止模型被重复加载）
_predictor_instance = None
_predictor_lock = threading.Lock()


def get_predictor():
    """获取预测器单例（线程安全）"""
    global _predictor_instance
    if _predictor_instance is None:
        with _predictor_lock:
            if _predictor_instance is None:
                try:
                    _predictor_instance = COSMICModelPredictor()
                except Exception as e:
                    log_warn(f"⚠️ 无法初始化模型预测器: {str(e)}")

                    log_warn("⚠️ 将使用基于规则的判断方法")
                    _predictor_instance = None
    return _predictor_instance
