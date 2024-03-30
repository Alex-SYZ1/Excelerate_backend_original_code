import zipfile
import io,os
from method2_1_tmp_to_zip import open_folder_choose_file


if __name__ == "__main__":
    # 模拟后端处理完的数据
    file_data = {
        "1.txt": io.BytesIO(b"Example content for file 1."),
        "2.txt": io.BytesIO(b"Example content for file 2."),
        "3.txt": io.BytesIO(b"Example content for file 3.")
    }
    # 将相对路径改到本文件夹
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    # 创建Zip文件的io.BytesIO对象
    zip_buffer = io.BytesIO()

    # 将文件数据直接写入到zip文件
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for filename, data in file_data.items():
            data.seek(0)
            zip_file.writestr(filename, data.read())

    # 准备好io.BytesIO对象以供下载
    zip_buffer.seek(0)
    file_path=os.path.join(os.getcwd(),"method2-2直接压缩后.zip")
    # 这里仅仅是为了演示，实际上应该是将zip_buffer传送给前端进行下载
    with open(file_path, "wb") as f:
        f.write(zip_buffer.getvalue())
    print("Zip created from data streams.",zip_buffer)

    open_folder_choose_file(file_path)