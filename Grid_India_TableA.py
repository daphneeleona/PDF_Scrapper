import streamlit as st

"""
## Grid India PSP Report Extractor (Chrome + Streamlit)

This app scrapes PSP Excel reports from Grid India using Selenium with Chrome (Chromium).
"""

with st.echo():
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

    @st.cache_resource
    def get_driver():
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1920,1080")

        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

    # --- Streamlit UI ---
    st.title("Grid India PSP Report Extractor")

    years = [f"{y}-{str(y+1)[-2:]}" for y in range(2013, 2026)]
    selected_year = st.selectbox("Select Financial Year", years[::-1])

    months = ["ALL", "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
    selected_month = st.selectbox("Select Month", months)

    if st.button("Extract Data"):
        with st.spinner("Scraping data... Please wait."):
            driver = get_driver()
            driver.get("https://grid-india.in/en/reports/daily-psp-report")
            wait = WebDriverWait(driver, 30)

            # Select year
            dropdown1 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp .my-select__control")))
            dropdown1.click()
            option1 = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_year}')]")))
            option1.click()

            # Select month
            dropdown2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".period_drp.me-1 .my-select__control")))
            dropdown2.click()
            option2 = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{selected_month}')]")))
            option2.click()

            time.sleep(10)

            # Show 100 entries
            Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Choose a page size']")))).select_by_visible_text("100")
            time.sleep(10)

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
                                    if report_date.month in [4, 5, 6] and 1 <= report_date.day <= 31:
                                        excel_links.append((report_date, href))
                                except:
                                    continue
                except Exception as e:
                    st.error(f"Error locating or reading the table: {e}")

            extract_links_from_table()

            try:
                next_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Next Page']")))
                if next_button.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    next_button.click()
                    time.sleep(5)
                    extract_links_from_table()
            except Exception as e:
                st.warning(f"Error checking or clicking 'Next Page': {e}")

            driver.quit()

            # Process Excel files
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
                st.success("Data extraction complete!")

                st.download_button(
                    label="📥 Download Excel",
                    data=output.getvalue(),
                    file_name="tut_purpose.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No data extracted.")
