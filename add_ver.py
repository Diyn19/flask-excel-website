import openpyxl
from datetime import datetime
import threading
import sys

def get_version(timeout=10, auto_only=False):
    version = None
    user_input = {"value": None}

    if not auto_only:
        def ask_input():
            try:
                user_input["value"] = input(f"請輸入版本號（{timeout} 秒內輸入，否則自動填入 MMDDHHMM）：\n> ").strip()
            except EOFError:
                user_input["value"] = ""

        t = threading.Thread(target=ask_input)
        t.daemon = True
        t.start()
        t.join(timeout=timeout)

        if user_input["value"]:
            version = user_input["value"]

    if not version:
        version = datetime.now().strftime("%m%d%H%M")
        if not auto_only:
            print(f"\n⏰ 超過 {timeout} 秒未輸入，自動使用版本號 {version}")

    # 寫入 Excel
    file_path = "data.xlsx"
    wb = openpyxl.load_workbook(file_path)
    sheet = wb["首頁"]
    sheet["G1"] = version
    wb.save(file_path)

    print(f"[完成] 已將版本號 {version} 寫入 G1")
    # 輸出給 BAT
    print(version)
    return version

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--auto", action="store_true", help="自動生成版本號（Git-only 模式）")
    args = parser.parse_args()
    get_version(auto_only=args.auto)
