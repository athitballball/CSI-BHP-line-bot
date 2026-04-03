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

USERNAME     = os.environ[“CSI_USERNAME”]
PASSWORD     = os.environ[“CSI_PASSWORD”]
GOOGLE_CREDS = os.environ[“GOOGLE_CREDENTIALS”]
SHEET_ID     = “11dX9ga5X5yZBeL-Nb__F1bIS96QdIVbPZJ93QX7e0_E”
LOGIN_URL    = “https://csi-bdms-mgrs.azurewebsites.net”

# ── วันที่: เมื่อวาน → วันนี้ ──────────────────────────────────────────────────

yesterday = datetime.now() - timedelta(days=1)
today     = datetime.now()
START_DATE = yesterday.strftime(”%d/%b/%Y”)   # เช่น 02/Apr/2026
END_DATE   = today.strftime(”%d/%b/%Y”)       # เช่น 03/Apr/2026

print(f”Date range: {START_DATE} → {END_DATE}”)

def export_excel():
download_dir = “/tmp/downloads”
os.makedirs(download_dir, exist_ok=True)

```
# ล้างไฟล์เก่าใน download_dir ก่อนดาวน์โหลดใหม่
for old_file in glob.glob(download_dir + "/*.xlsx"):
    os.remove(old_file)

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
driver.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {"behavior": "allow", "downloadPath": download_dir}
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

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
    Select(driver.find_element(By.TAG_NAME, "select")).select_by_visible_text("BHP")
    print("Selected BHP")
    time.sleep(2)

    date_inputs = driver.find_elements(
        By.XPATH,
        "//label[contains(text(),'วันที่เริ่มต้น')]/..//input | //label[contains(text(),'วันที่สิ้นสุด')]/..//input"
    )
    print("Date inputs found: " + str(len(date_inputs)))

    if len(date_inputs) < 2:
        all_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
        print("All text inputs: " + str(len(all_inputs)))
        for i, inp in enumerate(all_inputs):
            print("input " + str(i) + " value: " + inp.get_attribute("value") + " id: " + str(inp.get_attribute("id")))
        date_inputs = all_inputs[1:3]

    # ── วันเริ่มต้น (เมื่อวาน) ──
    ActionChains(driver).move_to_element(date_inputs[0]).click().perform()
    time.sleep(1)
    date_inputs[0].send_keys(Keys.CONTROL + "a")
    date_inputs[0].send_keys(START_DATE)
    date_inputs[0].send_keys(Keys.ENTER)
    time.sleep(2)
    print("Start date typed: " + START_DATE)

    # ── วันสิ้นสุด (วันนี้) ──
    ActionChains(driver).move_to_element(date_inputs[1]).click().perform()
    time.sleep(1)
    date_inputs[1].send_keys(Keys.CONTROL + "a")
    date_inputs[1].send_keys(END_DATE)
    date_inputs[1].send_keys(Keys.ENTER)
    time.sleep(10)
    print("End date typed: " + END_DATE)

    driver.save_screenshot("/tmp/screenshot.png")
    print("Screenshot saved")

    labels = driver.find_elements(By.CSS_SELECTOR, "label")
    print("Found " + str(len(labels)) + " labels")
    for label in labels:
        text = label.text.strip()
        print("label: " + text)
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

    for i in range(120):
        files = glob.glob(download_dir + "/*.xlsx")
        complete = [f for f in files if not f.endswith(".crdownload")]
        if complete:
            filepath = max(complete, key=os.path.getctime)
            print("Downloaded: " + filepath)
            return filepath
        print(f"Waiting for file... ({i+1}/120)")
        time.sleep(1)

    print("No file found")
    return None

finally:
    driver.quit()
```

def upload_to_sheets(filepath):
creds = Credentials.from_service_account_info(
json.loads(GOOGLE_CREDS),
scopes=[“https://www.googleapis.com/auth/spreadsheets”]
)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)
wb = openpyxl.load_workbook(filepath)

```
for sheet_name in wb.sheetnames:
    ws_excel = wb[sheet_name]

    # อ่านข้อมูลจาก Excel ที่ดาวน์โหลดมา
    new_rows = [
        [str(cell) if cell is not None else "" for cell in row]
        for row in ws_excel.iter_rows(values_only=True)
    ]

    if not new_rows:
        print(f"Sheet {sheet_name}: ไม่มีข้อมูล ข้ามไป")
        continue

    header = new_rows[0]
    new_data_rows = new_rows[1:]  # ไม่รวม header

    # ── เปิด/สร้าง worksheet ใน Google Sheet ──
    try:
        worksheet = sh.worksheet(sheet_name)
        existing = worksheet.get_all_values()
    except Exception:
        worksheet = sh.add_worksheet(title=sheet_name, rows=5000, cols=30)
        existing = []

    if not existing:
        # ยังไม่มีข้อมูลเลย → เขียน header + data ทั้งหมด
        worksheet.update([header] + new_data_rows)
        print(f"Sheet '{sheet_name}': เขียนใหม่ {len(new_data_rows)} แถว (รวม header)")
        continue

    # ── Dedup: แปลงแถวที่มีอยู่แล้วเป็น set ──
    existing_set = set(tuple(row) for row in existing[1:])  # ข้าม header

    # กรองเฉพาะแถวใหม่ที่ยังไม่มีใน sheet
    rows_to_add = [
        row for row in new_data_rows
        if tuple(row) not in existing_set
    ]

    if rows_to_add:
        # append ต่อท้ายข้อมูลเดิม
        worksheet.append_rows(rows_to_add, value_input_option="RAW")
        print(f"Sheet '{sheet_name}': เพิ่ม {len(rows_to_add)} แถวใหม่ (ข้าม {len(new_data_rows) - len(rows_to_add)} แถวซ้ำ)")
    else:
        print(f"Sheet '{sheet_name}': ไม่มีแถวใหม่ (ทุกแถวซ้ำกับที่มีอยู่แล้ว)")

print("Done: https://docs.google.com/spreadsheets/d/" + SHEET_ID)
```

filepath = export_excel()
if filepath:
upload_to_sheets(filepath)
else:
print(“Export failed”)