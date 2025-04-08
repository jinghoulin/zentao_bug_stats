import pandas as pd
from tkinter import Tk
from tkinter import filedialog

def select_excel_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def process_excel(file_path):
    # 保留的列名
    columns_to_keep = [
        "Bug编号", "Bug标题", "关键词", "严重程度", "Bug类型", "Bug状态",
        "指派给", "解决者", "解决日期", "相关Bug", "根因分析", "解决措施",
        "责任方", "发生频率", "子状态"
    ]
    
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 保留指定列，删除其他列
    df = df[columns_to_keep]
    
    # 保存到新的Excel文件
    output_file_path = "processed_" + file_path.split('/')[-1]
    df.to_excel(output_file_path, index=False)
    print(f"处理后的文件已保存为: {output_file_path}")

if __name__ == "__main__":
    file_path = select_excel_file()
    if file_path:
        process_excel(file_path)
    else:
        print("未选择文件")