import os

def get_all_filenames_in_directory(directory_path):
    filenames = os.listdir(directory_path)
    # 过滤掉目录，只保留文件名
    filenames = [filename for filename in filenames if os.path.isfile(os.path.join(directory_path, filename))]
    return filenames

if __name__ == "__main__":
    directory_path = r"C:\Users\7\Desktop\PPT\PPT模板"  # 指定要获取文件名的目录路径

    all_filenames = get_all_filenames_in_directory(directory_path)
    print(all_filenames)
