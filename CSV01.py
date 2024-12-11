import pandas as pd
import os

# 原始 Excel 檔案資料夾
input_folder = "C:\資產負債表"
# 處理後檔案儲存資料夾
output_folder = "C:\資產負債表\ProcessedExcelFiles"

# 如果輸出資料夾不存在，則建立
os.makedirs(output_folder, exist_ok=True)

# 列出所有 Excel 檔案
excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]

# 計算全局流水號
global_counter = 1

# 遍歷每個 Excel 檔案並處理
for file in excel_files:
    input_path = os.path.join(input_folder, file)
    
    # 嘗試讀取檔案
    try:
        # 讀取 Excel 檔案
        df = pd.read_excel(input_path)
        
        # 生成流水號欄名
        column_count = len(df.columns)
        new_columns = [f"Col_{global_counter + i}" for i in range(column_count)]
        global_counter += column_count
        
        # 替換欄名
        df.columns = new_columns
        
        # 儲存處理後的檔案
        output_path = os.path.join(output_folder, file)
        df.to_excel(output_path, index=False)
        
        print(f"已成功處理檔案：{file}，並儲存到：{output_path}")
    except Exception as e:
        print(f"處理檔案 {file} 時發生錯誤：{e}")

print("所有檔案處理完成！")

