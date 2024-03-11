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
**使用convert_to_json_stream将python对象转化为json数据流**，其中转化为json的步骤包括但不限于：
  - `dict`：会被转换为 JSON 对象。
  - `list`, `tuple`：会被转换为 JSON 数组。
  - `str`：会被转换为 JSON 字符串。
  - `int`, `float`, `int- & float-derived Enums`：会被转换为 JSON 数字。
  - `True`/`False`：会被转换为 JSON 的 `true`/`false`。
  - `None`：会被转换为 JSON 的 `null`。
+ **关于程序的warning信息**：
  + 为防止px打开xlsx时不必要的警告信息：
  ```python
  "openpyxl\worksheet\_reader.py:329: UserWarning: Data Validation extension is not supported and will be removed"
  warn(msg)  

  + 使用以下代码：
  warnings.filterwarnings("ignore", category=UserWarning)  
  ```
  + 警告原因：
    在使用`openpyxl`库处理Excel文件时，可以添加数据验证的规则来控制单元格中可以输入的数据类型。然而，`openpyxl`并不支持所有的Excel数据验证规则。截至我知识更新的时间点（2023年），`openpyxl`主要支持以下几种数据验证规则：

    1. **Whole number** (`whole`): 输入值必须是整数。
    2. **Decimal** (`decimal`): 输入值必须是小数。
    3. **List** (`list`): 输入值必须是列表中的项。
    4. **Date** (`date`): 输入值必须是日期。
    5. **Time** (`time`): 输入值必须是时间。
    6. **TextLength** (`textLength`): 文本的长度必须满足特定要求。
    7. **Custom** (`custom`): 定义自己的公式，输入值需要符合这个公式。

  不支持的数据验证规则可能包括但不限于：

    1. **Color Scale**（色阶）: 在单元格中根据其值显示不同的颜色。
    2. **Icon Sets**（图标集）: 根据单元格的值，显示不同的图标。
    3. **Data Bars**（数据条）: 在单元格中以数据条的形式显示值的大小。
    4. **Top 10 items**（前10项）、**Above or below average**（高于或低于平均值）、**Unique or Duplicate**（唯一或重复）等条件格式。

  需要注意的是，`openpyxl`中的数据验证主要是指对单元格中可以输入的数据类型的约束，而不是对单元格显示格式的控制，后者通常是通过**条件格式**（Conditional Formatting）来实现的。如果你需要使用到这些不支持的特性，你可能需要寻求其他库或者手动在Excel中设置这些规则。
  
  **一般条件格式用的不多，故忽视。**
+ excel_processor.py
    **若使用pywin32转化xlsx为xls，即在excel_processor.py运行以下代码**
    ```python
    Xio=Excel_IO()
    xlsx_stream=io.BytesIO()
    excel_got=r"tests\for_xls2xlsx\xlsx_file.xlsx"
    with open(excel_got, 'rb') as file:
        xlsx_stream.write(file.read())
    # 重置流的位置到开始处，这样就可以从头读取
    xlsx_stream.seek(0)

    Xio.convert_excel_format(xlsx_stream,"xlsx","xls",True)
    input("xlsx文件已转化为xls文件，保存在tests/for_xls2xlsx目录下，请查看")
    ```
    **pywin32**运行本地的excel程序会**弹框警告**如下图，  
    + 原因可能是**Microsoft Excel软件的问题**
    ![windows_Excel程序报错](docs\images\windows_Excel程序警告.png)
    这种“兼容性警告”与本地的Excel版本、系统无关，仅仅是xls与xlsx的版本功能迭代原因
    理由：
    + 当使用Excel版本为windows Excel2019版本(如下图)时，pywin32转化xlsx为xls会警告，手动转化xlsx为xls同样警告
    ![windows的Excel版本信息](docs\images\windows的Excel版本信息.png)
    + 使用Excel版本为mac Excel2024版本(如下图)时，手动转化xlsx为xls同样会警告
    ![mac_Excel程序报错](docs\images\mac_Excel程序警告.png)
    ![mac的Excel版本信息](docs\images\mac的Excel版本信息.png)
    + 然而若使用WPS手动打开excel，将xlsx另存为xls，并不会产生此错误，且**引用其他工作表中的值的数据验证规则**的功能在xls文件中仍然存在(可能WPS在底层读取xls时兼容了xlsx的功能)；再将WPS转化产生的xls用Excel打开，同样未产生问题
    ![WPS实现格式转化无警告](docs\images\WPS实现格式转化无警告.png)  
    
    因此，excel_processor.convert_excel_format在windows系统中能完成的功能：  
    + 收到xls文件(流)，无损转化为xlsx文件(即把旧版本转化为新版本)    
    + 收到xlsx文件(流)，有损转化为xls文件(即把新版转化为旧版,新版本的部分功能无法转给旧版本)     
    + 支持**收到xls文件(流)，无损转化为xlsx文件，再直接转化为xls文件(流)(即把旧版本转化为新版本)**(因为旧版本文件A1转为新版本文笔A2后，A2必定也只使用了A1使用了的旧版本功能，因此A2再转化为旧版本文件，不会报错。)但一般来说，文件转化为xlsx后，会使用xlsx独有的功能，尤其是**引用其他工作表中的值的数据验证规则**,与我们的程序紧密相关，因此这种完美状况不太现实。
    **由上**建议把程序的文件格式转化相关功能相关设置为：
      + 建议用户上传xlsx文件，这样有利于数据的处理。若为xls文件，建议转化为xlsx文件，但程序不会强制如此
      + 当用户上传xls文件时，提醒用户我们将转化为xlsx进行操作，不会造成数据丢失
      + 程序最终生成的文件以xlsx格式保存，若需保存为xls请自行手动转化格式。
      **这样就不会造成数据的丢失**
### `FileRuleMaker` (services/file_rule_maker.py)
+ 为便于理解，后端方法的parameters，returns信息均包括了格式和内容
+ 后端代码行中的：
  ```python
  ""#....
  "
  ```
  之间的内容是方便测试、理解的初步代码。  
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
+ 2024.2.28
  + 确定了学校端服务类(FileRuleMaker)下各个方法的参数、返回值，进行了详细的解释。暂未修改方法下对应代码，前端可根据参数、返回值、解释理解后端方法，先行进行前端代码编写。
+ 2024.2.29-3.2
  + 去掉学校端服务类(FileRuleMaker)下第一方法，修改各方法参数、返回值类型，研究xls2xlsx
+ 2024.3.2
  + 添加学校端服务类(FileRuleMaker)下的get_file方法，研究成功规则写入excel下拉列表
+ 2024.3.3
  + 添加学校端服务类(FileRuleMaker)下的create_final_rules_and_examples方法，完善之前方法少数细节
+ 2024.3.4
  + 实现了convert_excel_format文件格式转化方法
  + 添加学校端服务类(FileRuleMaker)下的下列方法
    + extract_fields_from_excel
    + generate_user_rule_dict
    + create_final_rules_and_examples  
  + 并在main函数中将三个方法详细地解释、连贯地运行   
  事实上剩下的save_final_rules需要的工作较少了，目前已经基本涉及了其功能实现，后续视前端进度再完善。
  + 在README文件中，对xls2xlsx库、Excel文件格式转化的问题进行了探讨，并初步设想了convert_excel_format在程序中的应用方式。
+ 2024.3.7
  + 基本实现了file_rule_maker,顺畅运行其下的四个方法
+ 2024.3.10
  + 更新规则dict 含数据起始行
+ 2024.3.11
  + file_validator文件的读取、检验、保存三个方法均实现并顺畅运行，辅以示例、注释
## 联系方式
|姓名|电话|微信号|邮箱|
|---|---|---|---|
|税远志|18511682594|Avid825|2100016640@stu.pku.edu.cn||
