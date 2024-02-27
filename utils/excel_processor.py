import openpyxl as px
import io,json,re
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment, Protection
from openpyxl.utils import range_boundaries



class Excel_IO:
    
    def read_excel_file(self, excel_path, sheet_index=0):
        """openpyxl读取某路径的excel文件"""
        try:
            excel_wb = px.load_workbook(excel_path, data_only=True)
            excel_ws = excel_wb.worksheets[sheet_index]
            return (excel_wb,excel_ws)
        except IOError as e:
            print(f"An error occurred during reading: {e}")
            # Handle the exception as needed
            return None

    def load_workbook_from_stream(self,excel_stream, sheet_index=0):
        """openpyxl读取某数据流的excel文件"""
        if 1:#try:
            # 读取流中的内容为二进制数据
            excel_data = excel_stream.read()
            # 使用BytesIO创建一个类似文件的对象
            excel_bytes = io.BytesIO(excel_data)
            return self.read_excel_file(excel_bytes)
            """excel_wb = px.load_workbook(filename=excel_bytes,data_only=True)
            excel_ws = excel_wb.worksheets[sheet_index]
            return (excel_wb,excel_ws)
        except IOError as e:
            print(f"An error occurred: {e}")
            return None"""

    def save_excel(self, excel_wb, excel_path):
        """openpyxl传输wb对象到excel文件"""
        try:
            excel_wb.save(excel_path)
        except IOError as e:
            print(f"An error occurred during saving: {e}")

            
    def stream_excel_to_frontend(self, excel_wb):
        """openpyxl传输wb对象到excel数据流"""
        try:
            # 创建一个BytesIO对象来保存Excel文件
            excel_stream = io.BytesIO()
            
            # 将workbook保存到这个BytesIO流中
            excel_wb.save(excel_stream)

            # 重置流的位置到开始处，这样就可以读取它的内容了
            excel_stream.seek(0)

            # 返回流对象
            return excel_stream
        except IOError as e:
            print(f"An error occurred during streaming: {e}")
            return None

    
    
class Excel_attribute:
    """目前只考虑了一个工作簿&其一个工作表的修改，QUESTION无法实现多个工作表同时修改"""
    def __init__(self, excel_wb=None , excel_ws=None):
        """类无传输值分别表示创建新wb、读取wb第一个工作表"""
        if excel_wb is None:
            self.excel_wb = px.Workbook()
            self.excel_ws = self.excel_wb.active
        else:
            self.excel_wb = excel_wb
            self.excel_ws = excel_ws if excel_ws is not None else excel_wb.worksheets[0]
    
    def get_some_axis_cells(self,index,value_only=True):
        """获取某一行/列的单元格，依据参数返回单元格对象或值的list"""
        transform_cell=lambda cell:cell.value if value_only==True else cell
        excel_field=[transform_cell(cell) for cell in self.excel_ws[index] if cell.value]
        return excel_field
    
    def get_max_row_col(self):
        """worksheet提供的属性来获取最大行列数问题：目前发现单元格有颜色填充、字色等也会被视为有内容的单元格；
           根据值遍历出的最大行列数则无此问题
           此外，纯下拉列表无选择值，二者都不会视为单元格有内容
           故返回两种方法分别产生的最大行列数集合"""
        px_max_row = self.excel_ws.max_row
        px_max_col = self.excel_ws.max_column
        value_max_row = 0
        value_max_col = 0
        for row in self.excel_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    value_max_row = max(value_max_row, cell.row)
                    value_max_col = max(value_max_col, cell.column)
        return {"max_row":{px_max_col,value_max_col},
                "max_col":{px_max_row,value_max_row}}

    def modify_cell_style(self, cell, font=None, border=None, fill=None, 
                          number_format=None, protection=None, 
                          hyperlink=None, alignment=None):
        """根据对象性参数修改某一单元格的字体、边框、填充、数字格式、保护方式、超文本、对齐格式等"""
        # Check if cell is a string reference or a Cell object
        if isinstance(cell, str):
            cell = self.excel_ws[cell]
            
        # Apply styles as provided
        # Create a dictionary with attribute names and the values provided
        style_attributes = {
            'font': font,'border': border,'alignment': alignment,
            'fill': fill,'number_format': number_format,
            'protection': protection,'hyperlink': hyperlink }

        # Apply styles as provided
        for attr_name, attr_value in style_attributes.items():
            if attr_value is not None:
                setattr(cell, attr_name, attr_value)  # Set attribute by name
    
    def modify_CertainRange_style(self, cell_range, **style_kwargs):
        """根据对象性参数修改某一单元格区域的字体、边框、填充、数字格式、保护方式、超文本、对齐格式等"""
        # Convert cell range to actual range
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        
        # Iterate over all cells in the range
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = self.excel_ws.cell(row=row, column=col)
                self.modify_cell_style(cell, **style_kwargs)
                
    def modify_MutipleRange_style(self, cell_range_list, **style_kwargs):
        """根据对象性参数修改某一多分布单元格区域的字体、边框、填充、数字格式、保护方式、超文本、对齐格式等"""
        if type(cell_range_list)== str:cell_range_list=[cell_range_list]
        for cell_range in cell_range_list:
            self.modify_CertainRange_style(cell_range,**style_kwargs)

    
    def get_dropdowns(self):
        """获取工作表内的各列的下拉列表字典，同一行多种下拉列表的以list组织"""
        def get_dropdowns_values(validation):
            result=validation.formula1
            
            # 进一步，下拉列表不仅仅为序列
            # 若值为工作表单元格引用
            
            #捕获组 (.*!)? 是可选的，用来匹配任意字符后跟一个感叹号 !，代表可能存在的工作表名称。
            pattern = r"^(.*!)?(\$?[A-Za-z]\$?\d+:\$?[A-Za-z]\$?\d+)$"
            match_=re.search(pattern,result)
            if match_:
                match_groups=match_.groups()
                # 若跨工作表引用(预计更为合理)
                if (match_groups)[0]:
                    dropdown_sourcesheet=self.excel_wb[match_groups[0][:-1]]
                else:dropdown_sourcesheet=self.excel_ws
                # 默认被引用为数据验证的单元格不止一个
                min_col, min_row, max_col, max_row = range_boundaries(match_groups[-1].replace('$', ''))
                value_list=[]
                for i in range(min_row, max_row+1):
                    for j in range(min_col, max_col+1):
                        value_list.append(dropdown_sourcesheet.cell(i, j).value)
                return value_list
            
            # 若值为简单的手动输入序列
            elif "," in result:
                # 去除首尾的引号后，直接拆分为值
                return result[1:-1].split(',')
            
        drop_row=dict()

        # 含有当前工作表的所有有效性验证的对象
        validations = self.excel_ws.data_validations.dataValidation
        for validation in validations:
            
            #当前有效性涉及区域
            cell=str(validation.sqref)
            
            #目前的方式，仅匹配下拉列表选择所以值的。进一步：考虑介于等多种方式
            result=(get_dropdowns_values(validation))

            #如果是多列的下拉列表相同，分别进行检验
            if " " in cell:
                cells=cell.split(" ")
                for i in cells:
                    if i[0] not in drop_row:drop_row[i[0]]=[result]
                    elif set(result) in [set(already_result) for already_result in drop_row[i[0]]]:continue
                    else:drop_row[i[0]].append(result)
            else:
                if (cell)[0] not in drop_row:drop_row[(cell)[0]]=[result]
                else:drop_row[cell[0]].append(result)
        return drop_row
def convert_to_json_stream(data):
    """将Python数据类型转化为JSON格式的字符串"""
    json_string = json.dumps(data)
    
    # 创建一个StringIO对象，它提供了文件类的接口
    json_stream = io.StringIO(json_string)
    
    # 返回数据流
    return json_stream

def read_from_json_stream(json_stream):
    """从JSON数据流中读取数据并转换为Python数据类型"""
    # 重置流的读取位置到起始处
    json_stream.seek(0)
    
    # 从数据流中读取JSON数据并转换为Python数据类型
    data = json.load(json_stream)
    
    # 返回Python数据类型
    return data


if "__main__" == __name__:
    excel_got=r"D:\SYZ_folder\活动\学工\23秋\A挑战杯-自动化\second_shot_0225-01-52\backend\tests\for_fuker.extract\test2_dropdown_hidensheet.xlsx"

    Xio=Excel_IO()
    excel_wb,excel_ws=Xio.load_workbook_from_stream(excel_got) if type(excel_got) !=str else Xio.read_excel_file(excel_got)
    excel_attr = Excel_attribute(excel_wb)  # 创建一个新的Excel工作簿
    for i,j in (excel_attr.get_dropdowns()).items():print("*",i,j)
    """excel_attr = Excel_attribute(None)  # 创建一个新的Excel工作簿
    # 设置一个单元格的样式
    excel_attr.modify_cell_style(
        'A1',
        font=Font(name='Calibri', size=12, bold=True),
        fill=PatternFill(fill_type='solid', start_color='FFFF00'),
        alignment=Alignment(horizontal='center', vertical='center')
    )
    # 设置一个多分布单元格区域的样式
    excel_attr.modify_MutipleRange_style(
        ["B1:C3","b4:C8"],
        font=Font(name='Calibri', size=12, bold=True),
        fill=PatternFill(fill_type='solid', start_color='569CD6'),
        alignment=Alignment(horizontal='center', vertical='center')
    )
    excel_attr.excel_wb.save('example.xlsx')"""
    



