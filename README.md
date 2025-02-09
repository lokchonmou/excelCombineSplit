# 合併及分發 Excel 工具 / Excel Combine and Split Tool

這是一個用python編寫的簡單工具，讓你可以方便地將excel合併或分發。

This project provides a simple way to combine and split Excel files using Python. Customize the scripts as needed and enjoy using the tool! 



## 1. 專案概述 / Project Overview

本專案提供兩個主要功能：
- **合併 Excel**：讀取指定資料夾中所有 `.xlsx` 檔案，將它們合併為一個 Excel 檔案，並輸出成格式化的檔案。
- **分發 Excel**：先將指定資料夾中所有 `.xlsx` 檔案合併成一個資料庫，再根據指定欄位（例如「產品名稱」）的值，分割成多個格式化的 Excel 檔案。

This project provides two main features:
- **Combine Excel Files**: Reads all `.xlsx` files from a specified folder, combines them into one Excel file, and outputs a formatted file.
- **Split Excel Files**: Combines all `.xlsx` files from a specified folder into a database first, then splits the data into multiple formatted Excel files based on a specified column (e.g., "產品名稱").

## 2. 檔案結構 / Project Structure
```
/project_folder
├── data # 原始 Excel 檔案存放處 / Folder storing original Excel files
├── output # 輸出結果存放處 / Folder to store output Excel files
├── `合併_分發.py` # 主程式，包含 combineExcel() 與 splitExcel() 函數 / Main script containing combineExcel() and splitExcel() functions
└── README.md # 說明文件 / This README file
```
## 3. 前置需求 / Prerequisites

- Python 3.x  
- 必要的 Python 套件: `pandas`, `openpyxl`  
  (請使用以下命令安裝所需套件)
  ```
  pip install pandas openpyxl
  ```

- Additionally, ensure that the module `prettyExcel` (included in your project) is available and that it contains the function `create_excel_with_format()` used for formatting Excel files.

## 4. 使用方法 / How to Use

### 4.1 合併 Excel 檔案 / Combining Excel Files

將所有需要合併的 `.xlsx` 檔案放入 `data` 資料夾中，然後執行下列程式碼範例（位於 `合併_分發.py` 中）：
```python
project_folder = os.getcwd()
data_folder = os.path.join(project_folder, "data")
output_folder = os.path.join(project_folder, "output")
combineExcel(data_folder, output_folder, 'combineExcel.xlsx')
```
執行後，合併後的格式化 Excel 檔案會儲存在 `output` 資料夾中，檔名為 `combineExcel.xlsx`。

Place all Excel files to combine in the `data` folder, then run the script as shown above. The combined and formatted Excel file will be output in the `output` folder.

### 4.2 分發 Excel 檔案 / Splitting Excel Files
假設 `data` 資料夾中已有一個或多個 Excel 檔案（或已合併的檔案），你可以根據指定的欄位（例如「產品名稱」）進行分發：
```python
project_folder = os.getcwd()
data_folder = os.path.join(project_folder, "data")
output_folder = os.path.join(project_folder, "output")
splitExcel(data_folder, output_folder, '產品名稱', output_filename='銷售資料')
```
這會根據「產品名稱」的不同值產生多個 Excel 檔案，檔名格式為 `銷售資料_產品名稱值.xlsx`，均儲存在 `output` 資料夾中。

If there is one or more Excel files (or a merged file) in the data folder, you can split the data based on a specific column (e.g., "產品名稱") as shown above. Multiple formatted Excel files will be generated in the `output` folder, with filenames like `銷售資料_<產品名稱值>.xlsx`.

## 5. 注意事項 / Notes
請確保 data 資料夾中的 Excel 檔案格式正確，且內含必要欄位名稱。

prettyExcel 模組中的 create_excel_with_format() 函數需可正常運作，才能正確格式化 Excel 檔案。
修改主程式 合併_分發.py 以符合您實際應用需求。

Make sure all Excel files in the data folder are correctly formatted and contain the expected columns. The create_excel_with_format() function from the prettyExcel module must work properly to format the Excel files according to the desired style. Adjust the main script as needed for your specific use case.

