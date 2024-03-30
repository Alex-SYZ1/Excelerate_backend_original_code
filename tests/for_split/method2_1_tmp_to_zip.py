import os
import zipfile
import io,time
from method1_tk_filedialog import open_folder
def open_folder_choose_file(file_path):
    import os
    import subprocess
    
    # 检查文件是否存在
    if os.path.exists(file_path):
        # 打开文件管理器并选中文件
        subprocess.run(f'explorer /select,"{file_path}"', shell=True)
        time.sleep(2)
        print("打开了生成文件的文件夹，并选中了该文件")
    else:
        print(f'文件 {file_path} 不存在。')
        
if __name__ == "__main__":
    # 模拟后端处理完的数据
    file_data = {
        "1.txt": io.BytesIO(b"Example content for file 1."),
        "2.txt": io.BytesIO(b"Example content for file 2."),
        "3.txt": io.BytesIO(b"Example content for file 3.")
    }
    # 将相对路径改到本文件夹
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    # 创建临时目录存放文件
    tmp_dir = "tmp/result_for_split"
    os.makedirs(tmp_dir, exist_ok=True)
    print("tmp" in os.listdir("."))
    print((os.getcwd()))
    # 保存数据流到临时文件
    for filename, file_io in file_data.items():
        with open(f"{tmp_dir}/{filename}", "wb") as file:
            file.write(file_io.getvalue())

    # 压缩临时文件夹
    zip_filename = "method2-1拆分后.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, dirs, files in os.walk(tmp_dir):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), os.path.join(tmp_dir, '..')))
    print(f"Folder compressed to {zip_filename}.")
    open_folder(os.path.join(os.getcwd(),tmp_dir))
    open_folder_choose_file(os.path.join(os.getcwd(),zip_filename))
