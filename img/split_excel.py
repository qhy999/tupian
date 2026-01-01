import pandas as pd
from datetime import datetime
import os

def split_excel_by_source(file_path):
    # 1. 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 '{file_path}'")
        return

    try:
        # 2. 读取 Excel 表格
        # 默认读取第一个工作表。如果不确定列名，我们通过索引（第1列和第2列）来操作
        df = pd.read_excel(file_path)
        
        # 获取列名（假设第一列是数据，第二列是来源）
        data_col = df.columns[0]
        source_col = df.columns[1]

        # 3. 获取当前日期 (格式为 MMDD)
        today_str = datetime.now().strftime("%m%d")

        # 4. 根据“来源”列进行分组处理
        grouped = df.groupby(source_col)

        for source_name, group_data in grouped:
            # 计算当前来源的数据条数
            count = len(group_data)
            
            # 5. 构建文件名：日期-来源-数量.xlsx
            # 注意：移除来源名称中可能导致文件名非法的字符
            clean_source_name = str(source_name).replace("/", "_").replace("\\", "_")
            new_filename = f"{today_str}-{clean_source_name}-{count}.xlsx"

            # 6. 保存到新的 Excel 文件
            group_data.to_excel(new_filename, index=False)
            print(f"已生成文件: {new_filename}")

        print("\n所有任务已完成！")

    except Exception as e:
        print(f"处理过程中出现错误: {e}")

if __name__ == "__main__":
    # 在这里输入你每天收到的原始文件名
    target_file = "今日数据.xlsx" 
    
    if os.path.exists(target_file):
        split_excel_by_source(target_file)
    else:
        # 如果找不到指定文件，提示用户手动输入
        manual_file = input("请输入 Excel 文件的完整名称（带后缀，如 input.xlsx）: ")
        split_excel_by_source(manual_file)