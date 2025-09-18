import openpyxl
from datetime import datetime
import threading

version = {"value": None}

def ask_input():
    try:
        version["value"] = input(
            "請輸入版本號（10 秒內輸入，否則自動填入 MMDDHHMM）："
        ).strip()
    except EOFError:
        version["value"] = ""

# 啟動輸入監聽執行緒
t = threading.Thread(target=ask_input)
t.daemon = True
t.start()
t.join(timeout=10)  # 最多等 10 秒

# 超時處理
if not version["value"]:
    version["value"] = datetime.now().strftime("%m%d%H%M")
    print(f"[超時] 10 秒未輸入，自動使用版本號 {version['value']}")

# 寫入 Excel
file_path = "data.xlsx"
wb = openpyxl.load_workbook(file_path)
if "首頁" not in wb.sheetnames:
    print("錯誤：找不到 '首頁' 工作表")
    exit(1)

sheet = wb["首頁"]
sheet["G1"] = version["value"]
wb.save(file_path)

print(f"✅ 已將版本號 {version['value']} 寫入 G1 儲存格")

# 輸出給批次檔
print(version["value"])
