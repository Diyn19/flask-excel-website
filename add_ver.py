import sys
import openpyxl
from datetime import datetime
import time
import msvcrt  # Windows 專用，可非阻塞檢測鍵盤輸入

version = None
timeout = 10  # 秒
start_time = time.time()

print(f"請輸入版本號（{timeout} 秒內輸入，否則自動填入 MMDDHHMM）：", end='', flush=True)
user_input = ''

while True:
    if msvcrt.kbhit():
        char = msvcrt.getwche()  # 讀取單個字元
        if char in ('\r', '\n'):
            break
        elif char == '\b':
            user_input = user_input[:-1]
            print('\b \b', end='', flush=True)
        else:
            user_input += char
    if time.time() - start_time > timeout:
        break
    time.sleep(0.01)

if user_input.strip():
    version = user_input.strip()
else:
    version = datetime.now().strftime("%m%d%H%M")
    print(f"\n[超時] 自動使用版本號 {version}")

# 寫入 Excel
file_path = "data.xlsx"
wb = openpyxl.load_workbook(file_path)
sheet = wb["首頁"]
sheet["G1"] = version
wb.save(file_path)

print(f"[完成] 已將版本號 {version} 寫入 G1")
# 輸出給 BAT
print(version)
