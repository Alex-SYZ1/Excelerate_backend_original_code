import difflib,re,random
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

def generate_strict_regex_and_example(input_list):
    # 使用 '^' 和 '$' 生成严格匹配列表中任一项的正则表达式
    regex_pattern = r'^(?:' + '|'.join(re.escape(item) for item in input_list) + r')$'
    
    # 从列表中随机选择一个样例
    random_example = random.choice(input_list)
    
    return [regex_pattern, random_example]

# 使用正则表达式
def match_with_regex(regex, string_to_test):
    return re.match(regex, string_to_test) is not None


"""# 示例使用match
options_list = ["填写", "日", "院系日名称", "院系"]
target_string = "日期"

# 输出匹配度最高的字符串
best_match_string = best_match(target_string, options_list)
print(best_match_string)
"""

if "__main__" == __name__:
    # 示例使用re
    my_list = ['apple', 'banana', 'cherry']
    regex, example = generate_strict_regex_and_example(my_list)

    # 验证正则表达式
    test_string = 'banana'
    if match_with_regex(regex, test_string):
        print(f'The string "{test_string}" is an exact match in the list.')
    else:
        print(f'The string "{test_string}" does not exactly match any item in the list.')

    print(f'Regex: {regex}')
    print(f'Random example: {example}')