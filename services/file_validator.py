import os,sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils.excel_processor

class FileValidator:

    def validate_filled_excel(self, filled_excel_file, rules_json):
        """
        验证已填写的Excel文件是否符合规则，并标记不符合的单元格。

        Parameters:
        filled_excel_file (file): 已填写的Excel文件。
        rules_json (json): 规则字典所在的JSON文件。

        Returns:
        tuple: (保存的Excel文件路径, 是否完全符合规则的标志)
        """
        

    def save_validated_excel(self, validated_excel_file, save_directory):
        """
        内容确认无误后，保存经过验证的Excel文件到本地目录。

        Parameters:
        validated_excel_file (file): 经过验证的Excel文件。
        save_directory (str): 文件保存目录。

        Returns:
        str: 保存的文件路径。
        """
        pass  # TODO: 实现方法