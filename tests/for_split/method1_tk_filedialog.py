import tkinter as tk
from tkinter import filedialog
import io,os,time
def open_folder(folder_path):
    import os
    import subprocess
    import sys
    if sys.platform == 'win32':  # Windows
        os.startfile(folder_path)
        time.sleep(1)
        print("打开了生成的文件夹")
    elif sys.platform == 'darwin':  # macOS
        subprocess.run(['open', folder_path])
    else:  # Linux and other Unix-like systems
        subprocess.run(['xdg-open', folder_path])


# 用于保存文件的函数
def save_files(file_data):
    root = tk.Tk()
    # 获取当前脚本的绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 构建 icon 的绝对路径
    icon_path = os.path.join(script_dir, "favicon.ico")

    # 设置窗口图标
    root.iconbitmap(icon_path)
    root.withdraw()  # 隐藏tkinter主窗口
    default_path = os.path.dirname(os.path.realpath(__file__))

    # 设置initialdir参数为代码所在文件夹
    folder_selected = filedialog.askdirectory(
        title="我们将保存文件夹至您选择的对应目录！(图标和文本都可修改)",
        initialdir=default_path
    )
    if folder_selected:
        for filename, file_io in file_data.items():
            file_path = f"{folder_selected}/{filename}"
            with open(file_path, "wb") as file:
                file.write(file_io.getvalue())
        print("Files saved successfully.")
    else:
        print("File save cancelled.")
    return folder_selected
if __name__ == "__main__":
    # 模拟后端处理完的数据
    file_data = {
        "1.txt": io.BytesIO(b"Example content for file 1."),
        "2.txt": io.BytesIO(b"Example content for file 2."),
        "3.txt": io.BytesIO(b"Example content for file 3.")
    }
    open_folder(save_files(file_data))
