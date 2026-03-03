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

USERNAME     = os.environ[“CSI_USERNAME”]
PASSWORD     = os.environ[“CSI_PASSWORD”]
GOOGLE_CREDS = os.environ[“GOOGLE_CREDENTIALS”]
SHEET_ID     = “11dX9ga5X5yZBeL-Nb__F1bIS96QdIVbPZJ93QX7e0_E”
LOGIN_URL    = “https://csi-bdms-mgrs.azurewebsites.net”
START_DATE   = “01/Mar/2026”

def export_excel():
download_dir = “/tmp/downloads”
os.makedirs(download_dir, exist_ok=True)

```
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
    print("Login success")

    # Export page
    driver.get(f"{LOGIN_URL}/Home/Export?uid=87")
    time.sleep(3)
    print("Export page loaded")

    # Set start date
    inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[type='text']")))
    start_input = inputs[0]
    driver.execute_script("arguments[0].click();", start_input)
    time.sleep(2)
    driver.save_screenshot("/tmp/screenshot.png")
    print("Screenshot saved")

    day_cells = driver.find_elements(By.CSS_SELECTOR, "td.available:not(.off)")
    print(f"Found {len(day_cells)} day cells")
    for cell in day_cells:
        if cell.text.strip() == "1":
            driver.execute_script("arguments[0].click();", cell)
            print("Clicked day 1")
            break
    time.sleep(0.5)

    apply_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".applyBtn")))
    driver.execute_script("arguments[0].click();", apply_btn)
    print("Start date set: " + START_DATE)
    time.sleep(1)

    # Set end date
    end_date = datetime.now().strftime("%d/%b/%Y")
    today_day = str(datetime.now().day)
    end_input = inputs[1]
    driver.execute_script("arguments[0].click();", end_input)
    time.sleep(2)

    day_cells = driver.find_elements(By.CSS_SELECTOR, "td.available:not(.off)")
    for cell in day_cells:
        if cell.text.strip() == today_day:
            driver.execute_script("arguments[0].click();", cell)
            print("Clicked day " + today_day)
            break
    time.sleep(0.5)

    apply_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".applyBtn")))
    driver.execute_script("arguments[0].click();", apply_btn)
    print("End date set: " + end_date)
    time.sleep(1)

    # Select Site = BHP
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
    Select(driver.find_element(By.TAG_NAME, "select")).select_by_visible_text("BHP")
    print("Selected BHP")
    time.sleep(2)

    # Tick checkboxes with numbers
    time.sleep(3)
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

    # Close date picker
    driver.execute_script("document.body.click();")
    time.sleep(1)

    # Click Export
    export_btn = wait.until(EC.presence_of_element_located((By.ID, "exportBtn")))
    driver.execute_script("arguments[0].click();", export_btn)
    print("Clicked Export")
    time.sleep(5)

    # Find downloaded file
    files = glob.glob(f"{download_dir}/*.xlsx") or glob.glob(f"{download_dir}/*")
    if files:
        filepath = max(files, key=os.path.getctime)
        print("Downloaded: " + filepath)
        return filepath
    else:
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
    ws = wb[sheet_name]
    data = [
        [str(cell) if cell is not None else "" for cell in row]
        for row in ws.iter_rows(values_only=True)
    ]
    try:
        worksheet = sh.worksheet(sheet_name)
        worksheet.clear()
    except Exception:
        worksheet = sh.add_worksheet(title=sheet_name, rows=5000, cols=30)

    worksheet.update(data)
    print("Updated sheet: " + sheet_name)

print("https://docs.google.com/spreadsheets/d/" + SHEET_ID)
```

# Main

filepath = export_excel()
if filepath:
upload_to_sheets(filepath)
else:
print(“Export failed”)
