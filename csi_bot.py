import requests
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

USERNAME   = os.environ["CSI_USERNAME"]
PASSWORD   = os.environ["CSI_PASSWORD"]
LINE_TOKEN = os.environ["LINE_TOKEN"]
LOGIN_URL  = "https://csi-bdms-mgrs.azurewebsites.net"

def scrape_csi():
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
driver = webdriver.Chrome(options=options)
    try:
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
        time.sleep(3)
        driver.get(f"{LOGIN_URL}/bhp/view-score")
        time.sleep(2)
        today = datetime.now().strftime("%d/%b/%Y")
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 2:
                data.append({"form": cols[0].text.strip(), "total": cols[1].text.strip()})
        return today, data
    finally:
        driver.quit()

def send_line(message):
    requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers={
            "Authorization": f"Bearer {LINE_TOKEN}",
            "Content-Type": "application/json"
        },
        json={
            "to": os.environ["LINE_GROUP_ID"],
            "messages": [{"type": "text", "text": message}]
        }
    )

def format_message(date, data):
    lines = [f"\n📊 CSI Daily Report", f"📅 {date}", "─"*25]
    for item in data:
        lines.append(f"📋 {item['form']}  →  {item['total']} รายการ")
    lines.append("─"*25)
    lines.append("🏥 Bangkok Hospital Pakchong")
    return "\n".join(lines)

today, data = scrape_csi()
message = format_message(today, data)
send_line(message)
print("✅ ส่งสำเร็จ")
