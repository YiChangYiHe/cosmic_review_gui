# utils/model_integration_patch.py
"""
COSMIC 评估模型集成补丁
用于将 AI 模型预测集成到现有评估流程中
"""

from utils.model_predictor import get_predictor
from utils.runtime_logger import log_info


def predict_with_model_or_rules(desc_val, proc_val, current_l3, 
                                entry_kws, read_kws, exit_kws, write_kws,
                                predictor=None):
    """
    使用AI模型预测或回退到规则引擎
    
    Args:
        desc_val: 子过程描述
        proc_val: 功能过程名称
        current_l3: 三级模块名称
        entry_kws, read_kws, exit_kws, write_kws: 规则关键词列表
        predictor: 模型预测器实例（可选）
        
    Returns:
        dict: {
            'ai_recognize': 预测的数据移动类型 (E/R/X/W/n),
            'all_ai_types': 所有可能的类型 (e.g., "E/X"),
            'matched_categories': 匹配的类别列表,
            'matched_kws': 匹配的关键词列表,
            'detail_reason': 详细原因,
            'prediction_source': 'model' or 'rule',
            'confidence': 置信度 (0-1)
        }
    """
    
    if predictor is None:
        predictor = get_predictor()
    
    # 如果模型可用，使用模型预测
    if predictor:
        function_name = str(proc_val or "")
        
        # 步骤1：判断是否利旧
        is_legacy, legacy_conf, legacy_reason = predictor.predict_legacy(
            description=desc_val,
            function_name=function_name,
            level3_module=current_l3
        )
        
        if is_legacy:
            return {
                'ai_recognize': 'n',
                'all_ai_types': 'n',
                'matched_categories': [],
                'matched_kws': [legacy_reason],
                'detail_reason': legacy_reason,
                'prediction_source': 'model',
                'confidence': legacy_conf
            }
        
        # 步骤2：预测数据移动类型
        move_type, move_conf, move_reason = predictor.predict_move_type(
            description=desc_val,
            function_name=function_name,
            level3_module=current_l3,
            is_legacy=False
        )
        
        # 将 move_type 拆分为列表 (e.g., "ERX" -> ['E', 'R', 'X'])
        matched_categories = list(move_type) if move_type != 'n' else []
        all_ai_types = '/'.join(matched_categories) if matched_categories else 'n'
        
        return {
            'ai_recognize': move_type[0] if move_type else 'E',  # 取第一个字符作为主要类型
            'all_ai_types': all_ai_types,
            'matched_categories': matched_categories,
            'matched_kws': [move_reason],
            'detail_reason': move_reason,
            'prediction_source': 'model',
            'confidence': move_conf
        }
    
    # 回退到规则引擎
    import re
    h_text = desc_val.strip().lower()
    h_text = re.sub(r"^[0-9.\-\s]+", "", h_text)
    
    matched_categories = []
    matched_kws = []
    
    # 规则匹配逻辑
    for kw in entry_kws:
        if kw in h_text:
            matched_categories.append("E")
            matched_kws.append(kw)
    
    for kw in read_kws:
        if kw in h_text:
            matched_categories.append("R")
            matched_kws.append(kw)
    
    for kw in exit_kws:
        if kw in h_text:
            matched_categories.append("X")
            matched_kws.append(kw)
    
    for kw in write_kws:
        if kw in h_text:
            matched_categories.append("W")
            matched_kws.append(kw)
    
    # 兜底模糊匹配
    if not matched_categories:
        for kw in ["请求", "提交", "选择", "勾选"]:
            if kw in h_text:
                matched_categories.append("E")
                matched_kws.append(kw)
        for kw in ["提示", "结果"]:
            if kw in h_text:
                matched_categories.append("X")
                matched_kws.append(kw)
    
    # 去重保留顺序
    matched_categories = list(dict.fromkeys(matched_categories))
    matched_kws = list(dict.fromkeys(matched_kws))
    
    all_ai_types = '/'.join(matched_categories) if matched_categories else 'n'
    ai_recognize = matched_categories[0] if matched_categories else 'n'
    
    # 生成原因
    if matched_kws:
        detail_reason = f"识别到关键字: {', '.join(matched_kws)}"
    else:
        detail_reason = "未匹配到关键字"
    
    return {
        'ai_recognize': ai_recognize,
        'all_ai_types': all_ai_types,
        'matched_categories': matched_categories,
        'matched_kws': matched_kws,
        'detail_reason': detail_reason,
        'prediction_source': 'rule',
        'confidence': 1.0
    }


# 使用示例
if __name__ == "__main__":
    # 测试集成
    test_cases = [
        "接收用户输入的查询条件",
        "读取缓存中的配置信息",
        "保存用户信息到数据库",
        "返回查询结果给前端",
        "校验输入参数格式"
    ]
    
    # 定义规则关键词
    entry_kws = ["接收", "送入", "输入", "点击", "发起"]
    read_kws = ["读取", "检索", "查询", "获取"]
    exit_kws = ["返回", "响应", "输出", "导出"]
    write_kws = ["写入", "更新", "保存", "存储", "修改"]
    
    log_info('=' * 60)
    log_info('AI模型预测测试')
    log_info('=' * 60)
    
    for desc in test_cases:
        result = predict_with_model_or_rules(
            desc_val=desc,
            proc_val="测试功能",
            current_l3="测试模块",
            entry_kws=entry_kws,
            read_kws=read_kws,
            exit_kws=exit_kws,
            write_kws=write_kws
        )
        
        log_info(f'\n描述: {desc}')
        log_info(f"预测类型: {result['ai_recognize']} (完整: {result['all_ai_types']})")
        log_info(f"原因: {result['detail_reason']}")
        log_info(f"来源: {result['prediction_source']}, 置信度: {result['confidence']:.2f}")
    
    log_info('\n' + '=' * 60)
