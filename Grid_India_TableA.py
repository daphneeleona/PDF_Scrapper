import streamlit as st
import time
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os

# ---------------------------
# ✅ Minimal, Streamlit-safe Driver Setup
# ---------------------------

options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")

# Explicitly point to Chromium binary (for Streamlit Cloud)
chrome_path = "/usr/bin/chromium"
if os.path.exists("/usr/bin/chromium-browser"):
    chrome_path = "/usr/bin/chromium-browser"
options.binary_location = chrome_path

@st.experimental_singleton
def get_driver():
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ---------------------------
# ✅ Streamlit UI
# ---------------------------

st.title("📊 Grid India PSP Report Extractor")

# Financial years limited to 2023–24 through 2025–26
years = ["2023-24", "2024-25", "2025-26"]
selected_year = st.selectbox("Select Financial Year", years[::-1])

months = ["ALL", "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
selected_month = st.selectbox("Select Month", months)

if st.button("🔍 Extract Data"):
    with st.spinner("Launching headless browser... Please wait."):
        driver = get_driver()
        driver.get("https://grid-india.in/en/reports/daily-psp-report")
        wait = WebDriverWait(driver, 30)

        # --- Select financial year ---
        dropdown1 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control")))
        dropdown1.click()
        option1 = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_year}')]")))
        option1.click()

        # --- Select month ---
        dropdown2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control")))
        dropdown2.click()
        option2 = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_month}')]")))
        option2.click()

        time.sleep(10)

        # --- Show 100 entries per page ---
        Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']")))).select_by_visible_text("100")
        time.sleep(5)

        excel_links = []

        def extract_links_from_table():
            try:
                table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/main/div/div[3]/div/div/div[2]/table')))
                rows = table.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    links = row.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        href = link.get_attribute("href")
                        if href and "PSP" in href and href.endswith((".xls", ".xlsx", ".XLS")):
                            try:
                                date_str = href.split("/")[-1].split("_")[0]
                                report_date = datetime.strptime(date_str, "%d.%m.%y")
                                excel_links.append((report_date, href))
                            except:
                                continue
            except Exception as e:
                st.error(f"Error locating or reading the table: {e}")

        # ✅ Extract all paginated results
        extract_links_from_table()
        while True:
            try:
                next_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Next Page']")))
                if not next_button.is_enabled():
                    break
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(1)
                next_button.click()
                time.sleep(5)
                extract_links_from_table()
            except Exception as e:
                st.warning(f"No more pages or pagination failed: {e}")
                break

        driver.quit()

        # ---------------------------
        # ✅ Download + Process Excel files
        # ---------------------------

        excel_links.sort(key=lambda x: x[0])
        expected_columns1 = ["Region", "NR", "WR", "SR", "ER", "NER", "Total", "Remarks"]
        table1_combined = []

        for report_date, url in excel_links:
            try:
                response = requests.get(url, verify=False)
                if response.status_code == 200:
                    ext = url.split(".")[-1].lower()
                    engine = "openpyxl" if ext == "xlsx" else "xlrd"
                    df_full = pd.read_excel(BytesIO(response.content), sheet_name="MOP_E", engine=engine, header=None)
                    df1 = df_full.iloc[5:13, :8].copy()
                    df1.columns = expected_columns1
                    df1.insert(0, "Date", report_date.strftime("%d-%m-%Y"))
                    table1_combined.append(df1)
            except Exception as e:
                st.warning(f"Failed to process {url}: {e}")

        if table1_combined:
            final_df = pd.concat(table1_combined, ignore_index=True)
            output = BytesIO()
            final_df.to_excel(output, index=False)
            st.success("✅ Data extraction complete!")

            st.download_button(
                label="📥 Download Excel",
                data=output.getvalue(),
                file_name="GridIndia_PSP_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ No data extracted.")
