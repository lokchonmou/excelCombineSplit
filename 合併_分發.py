import os
import pandas as pd
from prettyExcel import create_excel_with_format
from openpyxl import Workbook

def combineExcel(data_folder, output_folder, output_filename="combineExcel.xlsx"):
    '''
    合併 data_folder 內所有 .xlsx 檔案，並輸出為 output_folder/output_filename
    '''
    
    # 讀取 data_folder 內所有 .xlsx 檔案 (包含完整路徑)
    files = os.listdir(data_folder)
    xlsx_files = [os.path.join(data_folder, f) for f in files if f.endswith('.xlsx')]
    
    if not xlsx_files:
        print("資料夾內無任何 .xlsx 檔案")
        return

    # 讀取所有檔案並合併 DataFrame
    dataframes = [pd.read_excel(f) for f in xlsx_files]
    combined_df = pd.concat(dataframes, ignore_index=True)
    print(combined_df)
    # 確保 output_folder 存在，若無則建立
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 組合輸出檔案的完整路徑
    output_file = os.path.join(output_folder, output_filename)
    
    # 如果輸出檔案不存在，先建立一個空的工作簿，讓 prettyExcel 可以在其上操作
    if not os.path.exists(output_file):
        wb = Workbook()
        wb.save(output_file)
    
    # 使用 prettyExcel 將合併後的 dataframe 輸出為格式化 Excel
    create_excel_with_format(combined_df, output_file, sheet_name="Sheet", index=False, header=True)
    print("合併完成，檔案位置：", output_file)

def splitExcel(data_folder, output_folder, split_column, output_filename):
    '''
    依據 split_column 將 data_folder 內所有 .xlsx 檔案分割，並輸出至 output_folder, 檔名為 output_filename_{split_column}.xlsx
    '''
    # 讀取 data_folder 內所有 .xlsx 檔案 (包含完整路徑)
    files = os.listdir(data_folder)
    xlsx_files = [os.path.join(data_folder, f) for f in files if f.endswith('.xlsx')]
    
    if not xlsx_files:
        print("資料夾內無任何 .xlsx 檔案")
        return

    # 若有多個檔案，先合併成一個 database
    dataframes = [pd.read_excel(f) for f in xlsx_files]
    combined_df = pd.concat(dataframes, ignore_index=True)
    
    # 確保 output_folder 存在，若無則建立
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 根據 split_column 找出所有唯一值
    unique_values = combined_df[split_column].dropna().unique()
    
    for val in unique_values:
        # 過濾出符合 split_column 欄位值的資料
        subset_df = combined_df[combined_df[split_column] == val]
        
        # 用split_column作為檔名，替換掉可能影響檔案命名的字元
        safe_val = str(val).replace("/", "_")
        output_file = os.path.join(output_folder, f"{output_filename}_{safe_val}.xlsx")
        
        # 若 output_file 不存在，先建立一個空的工作簿讓 prettyExcel 可在其上操作
        if not os.path.exists(output_file):
            wb = Workbook()
            wb.save(output_file)
        
        # 用 prettyExcel 將分割後的 dataframe 輸出為格式化的 Excel 
        create_excel_with_format(subset_df, output_file, sheet_name="Sheet", index=False, header=True)
        print(f"已輸出 {output_file} (split_column: {val})")


if __name__ == "__main__":
    # 假設目前目錄就是 project folder
    project_folder = os.getcwd()
    data_folder = os.path.join(project_folder, "data")
    output_folder = os.path.join(project_folder, "output")
    # combineExcel(data_folder, output_folder, 'combineExcel.xlsx')
    splitExcel(data_folder, output_folder, '產品名稱', output_filename='銷售資料')