import os
import json
import glob
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials
import openpyxl

USERNAME     = os.environ["CSI_USERNAME"]
PASSWORD     = os.environ["CSI_PASSWORD"]
GOOGLE_CREDS = os.environ["GOOGLE_CREDENTIALS"]
SHEET_ID     = "11dX9ga5X5yZBeL-Nb__F1bIS96QdIVbPZJ93QX7e0_E"
LOGIN_URL    = "https://csi-bdms-mgrs.azurewebsites.net"

def export_excel():
    download_dir = "/tmp/downloads"
    os.makedirs(download_dir, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
    })

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    try:
        wait = WebDriverWait(driver, 20)

        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()
        wait.until(EC.url_contains("FirstPage"))
        print("Login success")

        driver.get(LOGIN_URL + "/Home/Export?uid=87")
        time.sleep(3)
        print("Export page loaded")

        # เลือก BHP ก่อน
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
        Select(driver.find_element(By.TAG_NAME, "select")).select_by_visible_text("BHP")
        print("Selected BHP")
        time.sleep(2)

        # วันเริ่มต้น = เมื่อวาน, วันสิ้นสุด = วันนี้
        start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%b/%Y")
        end_date   = datetime.now().strftime("%d/%b/%Y")

        # หา input วันที่
        date_inputs = driver.find_elements(By.XPATH, "//label[contains(text(),'วันที่เริ่มต้น')]/..//input | //label[contains(text(),'วันที่สิ้นสุด')]/..//input")
        if len(date_inputs) < 2:
            all_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
            date_inputs = all_inputs[1:3]

        # พิมพ์วันที่เริ่มต้น
        ActionChains(driver).move_to_element(date_inputs[0]).click().perform()
        time.sleep(1)
        date_inputs[0].send_keys(Keys.CONTROL + "a")
        date_inputs[0].send_keys(start_date)
        date_inputs[0].send_keys(Keys.ENTER)
        time.sleep(2)
        print("Start date typed: " + start_date)

        # พิมพ์วันที่สิ้นสุด
        ActionChains(driver).move_to_element(date_inputs[1]).click().perform()
        time.sleep(1)
        date_inputs[1].send_keys(Keys.CONTROL + "a")
        date_inputs[1].send_keys(end_date)
        date_inputs[1].send_keys(Keys.ENTER)
        time.sleep(10)
        print("End date typed: " + end_date)

        # ติ๊ก checkbox ที่มีตัวเลขในวงเล็บ
        labels = driver.find_elements(By.CSS_SELECTOR, "label")
        for label in labels:
            text = label.text.strip()
            if "(" in text and ")" in text:
                try:
                    checkbox_id = label.get_attribute("for")
                    if checkbox_id:
                        checkbox = driver.find_element(By.ID, checkbox_id)
                    else:
                        checkbox = label.find_element(By.TAG_NAME, "input")
                    if not checkbox.is_selected():
                        driver.execute_script("arguments[0].click();", checkbox)
                    print("Ticked: " + text)
                except Exception as e:
                    print("Could not tick: " + text + " -> " + str(e))
        time.sleep(1)

        driver.execute_script("document.body.click();")
        time.sleep(1)

        export_btn = wait.until(EC.presence_of_element_located((By.ID, "exportBtn")))
        driver.execute_script("arguments[0].click();", export_btn)
        print("Clicked Export")
        time.sleep(10)

        files = glob.glob(download_dir + "/*.xlsx") or glob.glob(download_dir + "/*")
        if files:
            filepath = max(files, key=os.path.getctime)
            print("Downloaded: " + filepath)
            return filepath
        else:
            print("No file found")
            return None

    finally:
        driver.quit()

def upload_to_sheets(filepath):
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDS),
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    wb = openpyxl.load_workbook(filepath)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        new_rows = [
            [str(cell) if cell is not None else "" for cell in row]
            for row in ws.iter_rows(values_only=True)
        ]
        if not new_rows:
            continue

        header = new_rows[0]
        data_rows = new_rows[1:]

        try:
            worksheet = sh.worksheet(sheet_name)
            existing = worksheet.get_all_values()

            # หา column index ของ cguid
            if existing:
                existing_header = existing[0]
                cguid_col = existing_header.index("cguid") if "cguid" in existing_header else 0
                existing_cguids = set(row[cguid_col] for row in existing[1:] if len(row) > cguid_col)
            else:
                cguid_col = 0
                existing_cguids = set()
                worksheet.append_row(header)

        except Exception:
            worksheet = sh.add_worksheet(title=sheet_name, rows=5000, cols=30)
            worksheet.append_row(header)
            existing_cguids = set()
            cguid_col = 0

        # กรอง row ที่ cguid ซ้ำออก
        new_header = header
        cguid_col_new = new_header.index("cguid") if "cguid" in new_header else 0
        added = 0
        for row in data_rows:
            cguid_val = row[cguid_col_new] if len(row) > cguid_col_new else ""
            if cguid_val and cguid_val not in existing_cguids:
                worksheet.append_row(row)
                existing_cguids.add(cguid_val)
                added += 1

        print("Sheet: " + sheet_name + " added " + str(added) + " new rows")

    print("https://docs.google.com/spreadsheets/d/" + SHEET_ID)

filepath = export_excel()
if filepath:
    upload_to_sheets(filepath)
else:
    print("Export failed")
