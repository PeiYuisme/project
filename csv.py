import pandas as pd
import os

# 指定存放 Excel 文件的資料夾路徑
folder_path = "C:\資產負債表\ProcessedExcelFiles"  # 替換為您的資料夾路徑
output_file = "merged_output.csv"  # 輸出的 CSV 文件名

# 初始化一個空的 DataFrame
merged_data = pd.DataFrame()

# 遍歷資料夾中的所有 Excel 文件
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):  # 過濾 Excel 文件
        file_path = os.path.join(folder_path, file_name)
        print(f"讀取文件: {file_path}")
        
        # 讀取 Excel 文件
        data = pd.read_excel(file_path)
        
        # 確保第一欄作為鍵進行合併
        if not merged_data.empty:
            merged_data = pd.merge(merged_data, data, on=data.columns[0], how='outer')
        else:
            merged_data = data

# 將合併的數據保存為 CSV 文件
merged_data.to_csv(output_file, index=False, encoding="utf-8")
print(f"所有 Excel 文件已根據第一欄合併並保存為: {output_file}")
