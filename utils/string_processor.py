import difflib
def have_common_characters(str1, str2):
    return bool(set(str1) & set(str2))
def best_match(target, options):
    """在选项中找到与目标字符串最接近的字符串。
    :param target: 目标字符串
    :param options: 字符串列表，用于与目标进行匹配
    :return: 匹配度最高的字符串
    """
    # 获取匹配度最高的字符串,
    matches = difflib.get_close_matches(target, options, n=1, cutoff=0.0)
    # 若没有或者甚至无共同字符，返回""
    if not matches:return ""
    elif not have_common_characters(target,matches[0]):return ""
    # 如果有匹配的，返回第一个（最佳匹配），否则返回None
    
    else:return matches[0]
# 示例使用
options_list = ["填写", "日", "院系日名称", "院系"]
target_string = "日期"

# 输出匹配度最高的字符串
best_match_string = best_match(target_string, options_list)
print(best_match_string)
