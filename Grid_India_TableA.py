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

st.set_page_config(page_title="Grid India PSP Report Extractor")

@st.cache_resource
def get_driver():
    options = Options()
    options.add_argument("--headless=new")  # Use new headless mode if available
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # You can add more options if needed

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

st.title("📊 Grid India PSP Report Extractor")

# Years: 2023-24 to 2025-26 only, reversed order for UI
years = [f"{y}-{str(y+1)[-2:]}" for y in range(2023, 2026)][::-1]
selected_year = st.selectbox("Select Financial Year", years)

months = ["ALL", "April", "May", "June", "July", "August", "September",
          "October", "November", "December", "January", "February", "March"]
selected_month = st.selectbox("Select Month", months)

if st.button("Extract Data"):
    with st.spinner("Scraping data... Please wait. This might take some time for large selections."):
        driver = get_driver()
        driver.get("https://grid-india.in/en/reports/daily-psp-report")
        wait = WebDriverWait(driver, 30)

        # Select year dropdown
        dropdown_year = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control")))
        dropdown_year.click()
        year_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_year}')]")))
        year_option.click()

        # Select month dropdown
        dropdown_month = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control")))
        dropdown_month.click()
        month_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_month}')]")))
        month_option.click()

        time.sleep(5)  # Wait for data to load

        # Show 100 entries per page
        page_size_select = Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']"))))
        page_size_select.select_by_visible_text("100")
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
                                # If user selected ALL months, accept all dates; else filter by selected month
                                if selected_month == "ALL" or report_date.strftime("%B") == selected_month:
                                    excel_links.append((report_date, href))
                            except Exception:
                                continue
            except Exception as e:
                st.error(f"Error locating or reading the table: {e}")

        # Extract from first page
        extract_links_from_table()

        # Pagination: keep clicking 'Next Page' and extracting links until disabled
        while True:
            try:
                next_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Next Page']")))
                if next_button.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    next_button.click()
                    time.sleep(5)  # wait for table to load
                    extract_links_from_table()
                else:
                    break
            except Exception:
                # No more pages or error, exit loop
                break

        driver.quit()

        if not excel_links:
            st.error("No PSP Excel report links found for the selected filters.")
        else:
            # Sort links by date ascending
            excel_links.sort(key=lambda x: x[0])

            expected_columns1 = ["Region", "NR", "WR", "SR", "ER", "NER", "Total", "Remarks"]
            table1_combined = []

            for report_date, url in excel_links:
                try:
                    response = requests.get(url, verify=False, timeout=30)
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
                output.seek(0)
                st.success(f"Data extraction complete! Extracted {len(table1_combined)} reports.")

                st.download_button(
                    label="📥 Download Excel",
