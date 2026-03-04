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

def set_date(driver, element, value):
driver.execute_script(”””
var el = arguments[0];
var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, ‘value’).set;
setter.call(el, arguments[1]);
el.dispatchEvent(new Event(‘input’, { bubbles: true }));
el.dispatchEvent(new Event(‘change’, { bubbles: true }));
“””, element, value)

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

    inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[type='text']")))

    set_date(driver, inputs[0], START_DATE)
    print("Start date set: " + START_DATE)
    time.sleep(0.5)

    end_date = datetime.now().strftime("%d/%b/%Y")
    set_date(driver, inputs[1], end_date)
    print("End date set: " + end_date)
    time.sleep(0.5)

    driver.save_screenshot("/tmp/screenshot.png")
    print("Screenshot saved")

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
    Select(driver.find_element(By.TAG_NAME, "select")).select_by_visible_text("BHP")
    print("Selected BHP")
    time.sleep(2)

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

    driver.execute_script("document.body.click();")
    time.sleep(1)

    export_btn = wait.until(EC.presence_of_element_located((By.ID, "exportBtn")))
    driver.execute_script("arguments[0].click();", export_btn)
    print("Clicked Export")
    time.sleep(5)

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

filepath = export_excel()
if filepath:
upload_to_sheets(filepath)
else:
print(“Export failed”)
