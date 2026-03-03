import os
import json
import glob
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials
import openpyxl

USERNAME     = os.environ["CSI_USERNAME"]
PASSWORD     = os.environ["CSI_PASSWORD"]
GOOGLE_CREDS = os.environ["GOOGLE_CREDENTIALS"]
SHEET_ID     = "11dX9ga5X5yZBeL-Nb__F1bIS96QdIVbPZJ93QX7e0_E"
LOGIN_URL    = "https://csi-bdms-mgrs.azurewebsites.net"
START_DATE   = "01/Mar/2026"

def export_excel():
    download_dir = "/tmp/downloads"
    os.makedirs(download_dir, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
    })

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    try:
        wait = WebDriverWait(driver, 20)

        # Login
        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()
        wait.until(EC.url_contains("FirstPage"))
        print("✅ Login สำเร็จ")

        # เข้าหน้า Export
        driver.get(f"{LOGIN_URL}/Home/Export?uid=87")
        time.sleep(3)
        print("✅ เข้าหน้า Export แล้ว")

        # Set วันที่เริ่มต้น
        inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[type='text']")))
        start_input = inputs[0]
        driver.execute_script("arguments[0].click();", start_input)
        time.sleep(2)
        driver.save_screenshot("/tmp/screenshot.png")
        print("✅ บันทึก screenshot แล้ว")

        # คลิกวันที่ 1 (เฉพาะที่อยู่ในเดือนปัจจุบัน ไม่ใช่ off)
        day_cells = driver.find_elements(By.CSS_SELECTOR, "td.available:not(.off)")
        print(f"พบ td.available {len(day_cells)} อัน")
        for cell in day_cells:
            if cell.text.strip() == "1":
                driver.execute_script("arguments[0].click();", cell)
                print("✅ คลิกวันที่ 1 แล้ว")
                break
        time.sleep(0.5)
        apply_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".​​​​​​​​​​​​​​​​
