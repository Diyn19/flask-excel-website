from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# 設定下載路徑
download_dir = r"D:\SynologyDrive\flask\IM"
file_name = "POS服務工作統計表.xls"
file_path = os.path.join(download_dir, file_name)

# 如果有同名檔案，先刪除
if os.path.exists(file_path):
    os.remove(file_path)
    print(f"已刪除舊檔：{file_path}")

# 設定 Edge 瀏覽器選項
edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True
edge_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# 啟動 Edge
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 15)

# 開啟網站
driver.get("http://192.168.1.252/service/")

# 登入頁面
wait.until(EC.presence_of_element_located((By.ID, "login_id"))).send_keys("pos0800")
driver.find_element(By.ID, "login_pwd").send_keys("Pos0800")
driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="登入"]').click()

# 切換 iframe（主內容區）
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframe")))

# 切回主框架，點選左側選單
driver.switch_to.default_content()
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
time.sleep(1)
driver.find_element(By.LINK_TEXT, "POS服務工作統計表").click()

# 回 iframe 中
driver.switch_to.frame("iframe")
time.sleep(2)

# 選擇「萊爾富」
driver.find_element(By.NAME, "customer").click()
time.sleep(0.5)
driver.find_element(By.XPATH, '//option[text()="萊爾富"]').click()

# 選擇「新北一部」
driver.find_element(By.NAME, "dept_id").click()
time.sleep(0.5)
driver.find_element(By.XPATH, '//option[text()="新北一部"]').click()

# 點擊查詢
driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
time.sleep(3)

# 匯出成EXCEL
driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

# 等待下載完成
time.sleep(10)

# 驗證檔案是否存在
if os.path.exists(file_path):
    print(f"下載成功：{file_path}")
else:
    print("⚠️ 找不到下載檔案")

# 結束
driver.quit()
