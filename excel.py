import pandas as pd
import os

# 获取所有Excel文件列表
def get_excel_files(directory):
    files = os.listdir(directory)
    excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xls')]
    return excel_files

def merge_excel_files(directory, output_filename):
    excel_files = get_excel_files(directory)
    
    # 初始化一个空的DataFrame用于存储合并后的数据
    data_frames = []

    # 遍历Excel文件列表，读取文件内容并合并
    for file in excel_files:
        file_path = os.path.join(directory, file)
        data = pd.read_excel(file_path)
        data_frames.append(data)

    # 使用pandas.concat合并数据
    merged_data = pd.concat(data_frames, ignore_index=True)
    

    # 将合并后的数据保存到新的Excel文件中
    merged_data.to_excel(output_filename, index=False)
    print(f"合并完成，已保存为 {output_filename}")

# 使用示例
input_directory = 'C:/Users/EDY/pro/excel'  # 包含多个Excel文件的文件夹
output_file = 'merged_data.xlsx'  # 合并后的Excel文件名
merge_excel_files(input_directory, output_file)
