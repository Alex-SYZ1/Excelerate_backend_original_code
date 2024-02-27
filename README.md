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

## 服务类方法说明

### `FileRuleMaker` (services/file_rule_maker.py)

- `extract_fields_from_excel`: 从Excel中提取字段行，并高亮显示。
- `generate_user_rule_dict`: 根据前端确认的字段和规则生成用户规则字典。
- `create_final_rules_and_examples`: 根据用户定义的规则生成最终的规则和示例。
- `save_final_rules`: 保存最终的规则和示例到JSON文件中。

### `FileValidator` (services/file_validator.py)

- `validate_filled_excel`: 验证已填写的Excel文件是否符合规则。
- `save_validated_excel`: 保存验证后的Excel文件到指定目录。

## 开发和部署

...

## 联系方式

...