import pandas as pd
import os

# 指定存放 Excel 文件的資料夾路徑
folder_path = "C:\資產負債表\ProcessedExcelFiles"  # 替換為您的資料夾路徑

# 指定處理後保存的資料夾
output_folder = "C:\資產負債表\ProcessedExcelFiles"
os.makedirs(output_folder, exist_ok=True)

# 遍歷資料夾中的所有 Excel 文件
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):  # 過濾 Excel 文件
        file_path = os.path.join(folder_path, file_name)
        print(f"正在處理文件: {file_path}")
        
        # 讀取 Excel 文件
        data = pd.read_excel(file_path)
        
        # 修改第一列的名稱
        if not data.empty:
            first_col_name = data.columns[0]
            data.rename(columns={first_col_name: "會計項目"}, inplace=True)
            print(f"第一列名稱由 '{first_col_name}' 修改為 '會計項目'")
        
        # 保存修改後的 Excel 文件
        output_path = os.path.join(output_folder, file_name)
        data.to_excel(output_path, index=False)
        print(f"已保存處理後的文件: {output_path}")
