import os
import time
import requests
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pdf2image import convert_from_bytes
import pytesseract

# Optional: Set path to Tesseract executable (Windows)
# pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# ✅ Set path to Poppler bin folder
poppler_path = r"C:\Users\\daphnee.dsouza\\Downloads\\poppler-25.07.0"  # Update this if your Poppler is in a different location

# Initialize WebDriver
edge_driver_path = "C:\\Users\\daphnee.dsouza\\Downloads\\edgedriver_win64\\msedgedriver.exe"
options = Options()
options.use_chromium = True
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Edge(options=options)
driver.get("https://grid-india.in/en/reports/daily-vre-report")
wait = WebDriverWait(driver, 30)

# Select year and ALL
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control"))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '2024-25')]"))).click()
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control"))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'ALL')]"))).click()
time.sleep(10)
Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']")))).select_by_visible_text("100")
time.sleep(10)

# Collect PDF links
pdf_links = []
for link in driver.find_elements(By.TAG_NAME, "a"):
    href = link.get_attribute("href")
    if href and "remc" in href.lower() and href.endswith(".pdf"):
        try:
            date_str = href.split("/")[-1].split("_")[0]
            report_date = datetime.strptime(date_str, "%d.%m.%Y")
            if report_date.month in [1, 2, 3]:
                pdf_links.append((report_date, href))
        except:
            continue

# OCR extraction function
def extract_table_8_ocr(pdf_url):
    try:
        response = requests.get(pdf_url, verify=False)
        response.raise_for_status()
        images = convert_from_bytes(response.content, dpi=300, poppler_path=poppler_path)
        table_8_text = ""
        found_table_8 = False
        for img in images:
            text = pytesseract.image_to_string(img)
            if "8." in text and "VRE Curtailment" in text:
                found_table_8 = True
            if found_table_8:
                table_8_text += text + "\n"
                if re.search(r"9\\.", text):
                    break
        return table_8_text.strip() if table_8_text else "Table Number 8 not found"
    except Exception as e:
        return f"Error: {e}"

# Extract and save
data = []
for report_date, pdf_url in pdf_links:
    print(f"Processing: {pdf_url}")
    content = extract_table_8_ocr(pdf_url)
    data.append({"Report Date": report_date.strftime("%d-%m-%Y"), "PDF Link": pdf_url, "Table 8 Content": content})

df = pd.DataFrame(data)
df.to_excel("table_8_ocr_output8.xlsx", index=False)
driver.quit()
print("✅ OCR extraction complete. Data saved to 'table_8_ocr_output8.xlsx'.")
