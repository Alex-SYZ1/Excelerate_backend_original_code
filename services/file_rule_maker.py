import os,sys
import openpyxl as px
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment, Protection

"""用于导入项目中不在同一文件夹的库"""
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import utils.excel_processor as XPRO

class FileRuleMaker:
    def __init__(self):
        pass
    def extract_fields_from_excel(self, excel_got):
        """
        接收Excel文件，读取第一个worksheet，寻找字段所在行，并提取所有字段值到列表。
        将字段行高亮，返回含字段名的列表和保存的Excel文件路径。

        Parameters:
        excel_file (file): 需要处理的Excel文件。

        Returns:
        tuple: (字段名的列表, 保存的Excel文件)
        """
        #获取文件并转化
        Xio=XPRO.Excel_IO()
        """项目实际部署时，无需判断是否为字符串，全部为前端发送的数据流"""
        excel_wb,excel_ws=Xio.load_workbook_from_stream(excel_got) if type(excel_got) !=str else Xio.read_excel_file(excel_got)
        
        #读取对象并获取属性
        Xattr=XPRO.Excel_attribute(excel_wb,excel_ws)
        All_rows=[Xattr.get_some_field_cells(index,value_only=False) for index in range(1,min(Xattr.get_max_row_col()["max_row"])+1)]
        
        #对预知的字段行进行操作
        Predict_fields_cells   = max(All_rows,key=len)
        Predict_fields_values  = [cell.value for cell in Predict_fields_cells]
        Predict_fields_indexes = [cell.coordinate for cell in Predict_fields_cells]
        Xattr.modify_MutipleRange_style(Predict_fields_indexes,fill=PatternFill(fill_type='solid', start_color='FFFF00'))
        
        """演示用，避免数据流无法理解"""
        excel_wb.save("../tests/output_test1.xlsx")
        
        OUTPUT1=XPRO.convert_to_json_stream(Predict_fields_values)
        OUTPUT2=Xio.stream_excel_to_frontend(excel_wb)
    
        return (OUTPUT1,OUTPUT2)
    
    def generate_user_rule_dict(self, excel_file, fields_list):
        """
        接收确认了字段行的Excel和字段名列表，输出包含预定义规则和下拉列表的字典到JSON。

        Parameters:
        excel_file (file): 确认了字段行的Excel文件。
        fields_list (list): 字段名列表。

        Returns:
        dict: 用户规则字典。
        """
        pass  # TODO: 实现方法

    def create_final_rules_and_examples(self, rules_json):
        """
        接收前端传来的规则字典所在的JSON文件，生成最终规则和样例。

        Parameters:
        rules_json (json): 规则字典所在的JSON文件。

        Returns:
        dict: 包含规则和样例的JSON对象。
        """
        pass  # TODO: 实现方法

    def save_final_rules(self, rules_dict, additional_info):
        """
        接收最终确认后的规则字典和其他提示信息，保存为JSON文件。

        Parameters:
        rules_dict (dict): 规则字典。
        additional_info (str): 其他提示信息。

        Returns:
        str: 保存的JSON文件路径。
        """
        pass  # TODO: 实现方法
    
if "__main__" == __name__:
    Fuker=FileRuleMaker()
    excel_got=r"../tests/test1.xlsx"
    print(Fuker.extract_fields_from_excel(excel_got))