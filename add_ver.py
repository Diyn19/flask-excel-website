import sys
import openpyxl

# 取得命令列參數：版本號
if len(sys.argv) != 2:
    print("請提供版本號作為參數，例如：python add_ver.py 123")
    sys.exit(1)

version = sys.argv[1]

# 讀取 Excel 檔案
file_path = "data.xlsx"
wb = openpyxl.load_workbook(file_path)

# 選取「首頁」工作表並寫入版本號到 G1
if "首頁" not in wb.sheetnames:
    print("錯誤：找不到 '首頁' 工作表")
    sys.exit(1)

sheet = wb["首頁"]
sheet["G1"] = version

# 儲存 Excel 檔案
wb.save(file_path)
print(f"已將版本號 {version} 寫入 G1 儲存格")
