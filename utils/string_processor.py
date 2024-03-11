import difflib,re,random,os
def get_filepath_variables(excel_got):
    """返回excel_got的folder_path,file_basename,file_extension,file_name的键值对字典"""
    folder_path=os.path.dirname(excel_got)
    file_basename,file_extension=os.path.splitext(os.path.basename(excel_got))
    file_name=os.path.basename(excel_got)
    return {name_:variable_ for name_,variable_ in zip("folder_path file_basename file_extension file_name".split(),[folder_path,file_basename,file_extension,file_name])}
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

def coordinate_from_string(coord_string):
    """Convert a coordinate string like 'B12' to a tuple ('B', 12)"""
    COORD_RE = re.compile(r'^[$]?([A-Za-z]{1,3})[$]?(\d+)$')
    match = COORD_RE.match(coord_string)
    if not match:
        msg = f"Invalid cell coordinates ({coord_string})"
        raise CellCoordinatesException(msg)
    column, row = match.groups()
    row = int(row)
    if not row:
        msg = f"There is no row 0 ({coord_string})"
        raise CellCoordinatesException(msg)
    return column, row

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