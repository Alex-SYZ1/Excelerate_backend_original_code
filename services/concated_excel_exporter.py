# for concat：
class ConcatedExcelExporter:
    def __init__(self, file_path_list,template_path): 
        """初始化方法，创建一个DataFrame和excel
        以及，"""
        self.file_path_list=file_path_list
    def send_template(self): 
        """传出样式模板表，前端选择行数列数位置"""
    def examine_file(self,file_to_examine=self.file_path_list):
        """检验表格格式，保留不符合原因，输出不符合路径列表"""
    def concat_files(self):
        """pd合并文件，px放入wb"""
        
    def set_style(self):
        """获取前端位置属性，自己生成样式属性并移植，return，前端保存"""
    
    