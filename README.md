# Excel内容收集：规则化内容部署&填写

该项目为Excel内容的规则化部署和填写提供后端支持。涉及的主要功能包括Excel文件读取、规则制定、内容验证等。

## 目录说明

- `services/`: 包含处理文件规则制定和验证的服务类。
- `utils/`: 包含辅助函数，如文件读取和写入。
- `models/`: 用于存放数据模型。
- `rules/`: 存放预设的规则文件。
- `resources/`: 包含非代码文件，如模板和配置文件。
- `tests/`: 包含项目的测试代码。
- `scripts/`: 包含项目启动和维护相关的脚本。
## 代码提醒事项
+ **前后端均可使用excel_processor.py中的read_from_json_stream方法读取数据流**
### `FileRuleMaker` (services/file_rule_maker.py)
+ 为便于理解，后端方法的parameters，returns信息均包括了格式和内容
+ 后端代码行中的
```python
""#
"
```
之间的内容为方便测试、理解的初步代码。
    + 若含"改为xxx" 可改为对应内容
    + 若含"可去除" 就可以直接去除
+ 后端暂未完善的功能：关于预定义规则的方法、内容(后续基本功能完成后，陆续添加。)
## 服务类方法说明


- `extract_fields_from_excel`: 从Excel中提取字段行，并高亮显示。(拟去除)
- `generate_user_rule_dict`: 根据前端确认的字段和规则生成用户规则字典。
- `create_final_rules_and_examples`: 根据用户定义的规则生成最终的规则和示例。
- `save_final_rules`: 保存最终的规则和示例到JSON文件中。

### `FileValidator` (services/file_validator.py)

- `validate_filled_excel`: 验证已填写的Excel文件是否符合规则。
- `save_validated_excel`: 保存验证后的Excel文件到指定目录。

## 更新日志
+ 2024.2.23  
    搭了整体框架，构造了几个略显优美的类及其方法
+ 2024.2.24  
  + 实现了学校端服务类(FileRuleMaker)下的与前端粗糙交互的第一个方法：获取excel文件数流，返回①保存字段名列表的json数据流；②高亮字段行的excel数据流。
    此方法尚以功能实现为主，未过于追求美观简化。
  + 解决项目内不同目录的库相互导入产生报错的问题：在每个py文件开头加上：
```python
import os,sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
```
即可 【以文件夹.文件.类/.】的方法 import package
  + 补充完善excel_processor下的类、方法，为之后的代码做准备
  + 导出虚拟环境Excelerate_backend至environment.yml文件
+ 2024.2.27
  + 实现了学校端服务类(FileRuleMaker)下的与前端粗糙交互的第二个方法generate_user_rule_dict，接收字段列表，匹配预定义&下拉规则与字段

## 联系方式
|姓名|电话|微信号|邮箱|
|---|---|---|---|
|税远志|18511682594|Avid825|2100016640@stu.pku.edu.cn||
