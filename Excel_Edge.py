from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import time
import shutil
import glob
from datetime import datetime

# 時間與檔名設定
yyyymm = datetime.now().strftime("%Y%m")
download_path = r"D:\flask\IM"
pos_pattern = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report*.xlsx")
mfp_pattern = os.path.join(download_path, f"{yyyymm}_Service_Count_Report*.xlsx")
pos_final = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
mfp_final = os.path.join(download_path, f"{yyyymm}_Service_Count_Report.xlsx")

# Edge 選項
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

try:
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 30)

    # 登入
    driver.get("http://eip.toshibatec.com.tw/Main.aspx")
    wait.until(EC.presence_of_element_located((By.NAME, "AccountID"))).send_keys("yang.di")
    driver.find_element(By.NAME, "PassWord").send_keys("foxdie789")
    driver.find_element(By.NAME, "login_SubmitBtn").click()

    # 點擊內部系統
    wait.until(EC.element_to_be_clickable((By.XPATH, '//td[text()="內部系統"]'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//td[contains(text(),"EIP 分析系統")]'))).click()
    time.sleep(3)

    # 切換到新開視窗
    driver.switch_to.window(driver.window_handles[-1])
    wait.until(EC.title_contains("台芝技術服務分析系統"))

    ### === POS服務工作統計表 === ###
    driver.switch_to.default_content()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "POS服務工作統計表"))).click()
    driver.switch_to.frame("iframe")
    time.sleep(2)

    for attempt in range(2):
        try:
            customer_field = wait.until(EC.element_to_be_clickable((By.NAME, "customer")))
            customer_field.click()
            time.sleep(0.5)
            customer_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//select[@name="customer"]/option[contains(text(),"萊爾富")]')))
            customer_option.click()

            dept_field = wait.until(EC.element_to_be_clickable((By.NAME, "dept_id")))
            dept_field.click()
            time.sleep(0.5)
            dept_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//select[@name="dept_id"]/option[contains(text(),"新北勤務一部")]')))
            dept_option.click()
            break
        except TimeoutException:
            if attempt == 0:
                print("⚠️ POS: 找不到 customer 或 dept_id，重新整理頁面...")
                driver.refresh()
                time.sleep(5)
                driver.switch_to.frame("iframe")
            else:
                print("❌ POS: 找不到 customer 或 dept_id，流程中止。")
                raise

    # 查詢並匯出
    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    for f in glob.glob(pos_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    downloaded_pos = None
    for _ in range(30):
        files = glob.glob(pos_pattern)
        if files:
            downloaded_pos = max(files, key=os.path.getctime)
            break
        time.sleep(1)

    if downloaded_pos and os.path.exists(downloaded_pos):
        shutil.move(downloaded_pos, pos_final)
        print(f"✅ POS報表下載完成：{os.path.basename(pos_final)}")
    else:
        print("❌ POS報表未下載完成")

    # 回上頁
    back_btn = wait.until(EC.element_to_be_clickable((By.ID, "back")))
    back_btn.click()
    time.sleep(2)

    ### === MFP服務工作統計表 === ###
    driver.switch_to.default_content()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "MFP服務工作統計表"))).click()
    driver.switch_to.frame("iframe")
    time.sleep(2)

    for attempt in range(2):
        try:
            dept_field = wait.until(EC.element_to_be_clickable((By.NAME, "dept_id")))
            dept_field.click()
            time.sleep(0.5)
            dept_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//select[@name="dept_id"]/option[contains(text(),"新北勤務一部")]')))
            dept_option.click()
            break
        except TimeoutException:
            if attempt == 0:
                print("⚠️ MFP: 找不到 dept_id，重新整理頁面...")
                driver.refresh()
                time.sleep(5)
                driver.switch_to.frame("iframe")
            else:
                print("❌ MFP: 找不到 dept_id，流程中止。")
                raise

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    for f in glob.glob(mfp_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    downloaded_mfp = None
    for _ in range(30):
        files = glob.glob(mfp_pattern)
        if files:
            downloaded_mfp = max(files, key=os.path.getctime)
            break
        time.sleep(1)

    if downloaded_mfp and os.path.exists(downloaded_mfp):
        shutil.move(downloaded_mfp, mfp_final)
        print(f"✅ MFP報表下載完成：{os.path.basename(mfp_final)}")
    else:
        print("❌ MFP報表未下載完成")

    driver.quit()

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()
