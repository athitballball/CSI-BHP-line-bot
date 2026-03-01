import requests
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

USERNAME   = os.environ["CSI_USERNAME"]
PASSWORD   = os.environ["CSI_PASSWORD"]
LINE_TOKEN = os.environ["LINE_TOKEN"]
LINE_GROUP_ID = os.environ["LINE_GROUP_ID"]

BASE_URL = "https://csi-bdms-mgrs.azurewebsites.net"
API_PATH = "/config/searchdashboard"

def get_session_cookies():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        driver.get(BASE_URL)
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))

        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()

        wait.until(EC.url_contains("/dashboard"))

        # get all cookies after login
        cookies = driver.get_cookies()
        session = requests.Session()
        for cookie in cookies:
            session.cookies.set(cookie["name"], cookie["value"])
        return session
    finally:
        driver.quit()

def call_dashboard_api(session, from_date, to_date, site_code):
    params = {
        "from": from_date,
        "to": to_date,
        "sitecode": site_code
    }
    url = BASE_URL + API_PATH
    resp = session.get(url, params=params)
    return resp.json()  # assuming JSON

def send_line(message):
    payload = {
        "to": LINE_GROUP_ID,
        "messages": [{"type": "text", "text": message}]
    }
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    requests.post("https://api.line.me/v2/bot/message/push", headers=headers, json=payload)

def format_message(data):
    today = datetime.now().strftime("%d/%b/%Y")
    lines = [f"📊 CSI Dashboard Report ({today})", "─────────────────"]
    for item in data:
        lines.append(f"{item['form']} → {item['total']}")
    return "\n".join(lines)

if __name__ == "__main__":
    session = get_session_cookies()
    api_data = call_dashboard_api(session, "01/Mar/2026", "01/Mar/2026", "BHP")

    # adjust depending on API structure
    if api_data:
        msg = format_message(api_data)
        send_line(msg)
        print("✅ ส่ง LINE สำเร็จ")
    else:
        print("❌ ไม่พบข้อมูล")
