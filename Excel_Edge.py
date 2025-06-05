from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil

# 設定 Edge 下載路徑
download_path = r"D:\SynologyDrive\flask\IM"
filename = "202506_HL_Maintain_Report.xlsx"
filepath = os.path.join(download_path, filename)

# Edge 選項
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# 啟動 Edge driver
driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 20)

# 開啟網站並登入
driver.get("http://192.168.1.252/service/")
wait.until(EC.presence_of_element_located((By.ID, "login_id"))).send_keys("pos0800")
driver.find_element(By.ID, "login_pwd").send_keys("Pos0800")
driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="登入"]').click()

# 切換至 iframe
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "iframe")))

# 回到主框架並操作主選單
driver.switch_to.default_content()
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
time.sleep(1)
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "POS服務工作統計表"))).click()

# 回到 iframe 等待查詢頁面載入
driver.switch_to.frame("iframe")
time.sleep(2)

# 點選「客戶」欄並選擇「萊爾富」
customer_field = wait.until(EC.element_to_be_clickable((By.NAME, "customer")))
customer_field.click()
time.sleep(0.5)
customer_option = wait.until(EC.element_to_be_clickable((By.XPATH, '//select[@name="customer"]/option[contains(text(),"萊爾富")]')))
customer_option.click()

# 點選「部門」欄並選擇「新北一部」
dept_field = wait.until(EC.element_to_be_clickable((By.NAME, "dept_id")))
dept_field.click()
time.sleep(0.5)
dept_option = wait.until(EC.element_to_be_clickable((By.XPATH, '//select[@name="dept_id"]/option[contains(text(),"新北勤務一部")]')))
dept_option.click()

# 按下查詢
driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()

# 等待「匯出成EXCEL」出現
wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

# 若檔案已存在則刪除（避免變成 (1).xlsx）
if os.path.exists(filepath):
    os.remove(filepath)

# 點擊「匯出成EXCEL」
driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

# 等待下載完成（可加強等待條件）
time.sleep(10)

# 關閉瀏覽器
driver.quit()
