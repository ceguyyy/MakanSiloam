import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import time
import sys
import os

# === FUNCTIONS ===
def get_weekdays(start_str, end_str):
    start = datetime.strptime(start_str, "%m/%d/%Y")
    end = datetime.strptime(end_str, "%m/%d/%Y")
    current = start
    weekdays = []
    while current <= end:
        if current.weekday() < 5:  # Weekday only
            weekdays.append(current.strftime("%m/%d/%Y"))
        current += timedelta(days=1)
    return weekdays

def setup_driver():
    """Set up Chrome WebDriver with headless options"""
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # Use WebDriver Manager to handle ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def automate_form(emp_status, nik, name, directorate_choice, dates):
    FORM_URL = "https://forms.office.com/pages/responsepage.aspx?id=UwNJH5TnM0e5oqJO5GV_T9KMRjpWHsZEkKVvBch9KKlUOFVBNFA1MEdMU0xKVFFZVjAzV0tPV1ZDSi4u&route=shorturl"

    status_xpath_map = {
        "Employee": '//input[@aria-label="Employee"]',
        "Vendor/Outsource": '//input[@aria-label="Vendor/Outsource"]',
        "Non Employee": '//input[@aria-label="Non Employee"]'
    }

    # Create a progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    driver = setup_driver()
    
    try:
        for i, date_str in enumerate(dates):
            progress = int((i + 1) / len(dates) * 100)
            progress_bar.progress(progress)
            status_text.text(f"Processing {i+1}/{len(dates)}: {date_str}...")
            
            driver.get(FORM_URL)
            time.sleep(2)  # Allow page to load

            # Select Employment Status
            driver.find_element(By.XPATH, status_xpath_map[emp_status]).click()
            time.sleep(0.5)

            # Input NIK
            nik_field = driver.find_element(By.XPATH, '//input[@aria-label="NIK"]')
            nik_field.clear()
            nik_field.send_keys(nik)
            time.sleep(0.5)

            # Input Name
            name_field = driver.find_element(By.XPATH, '//input[@aria-label="Nama Lengkap"]')
            name_field.clear()
            name_field.send_keys(name)
            time.sleep(0.5)

            # Select Directorate
            dir_xpath = f'//input[@aria-label="{directorate_choice}"]'
            driver.find_element(By.XPATH, dir_xpath).click()
            time.sleep(0.5)

            # Fill Date
            date_input = driver.find_element(By.XPATH, '//input[@aria-label="Date Picker"]')
            date_input.send_keys(Keys.CONTROL + "a")
            date_input.send_keys(Keys.DELETE)
            date_input.send_keys(date_str)
            date_input.send_keys(Keys.TAB)
            time.sleep(1)

            # Submit
            submit_btn = driver.find_element(By.XPATH, '//button[@data-automation-id="submitButton"]')
            submit_btn.click()
            time.sleep(2)  # Wait for submission

            st.success(f"✅ Submitted form for {date_str}")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.error(f"⚠️ Last processed date: {date_str}")
    finally:
        driver.quit()
        progress_bar.empty()
        status_text.empty()
        st.balloons()

# === STREAMLIT UI ===
st.title("🥪 Mempermudah Makan Anda di SHOO")
st.info("""
**Catatan Penggunaan:**
1. Form akan diisi secara otomatis untuk semua hari kerja dalam rentang tanggal
2. Pastikan NIK dan Nama Anda benar
3. Proses mungkin memakan waktu beberapa menit
""")

with st.expander("Form Input", expanded=True):
    emp_status = st.selectbox("Status Karyawan", ["Employee", "Vendor/Outsource", "Non Employee"])
    name = st.text_input("Nama Lengkap")
    nik = st.text_input("NIK (Khusus TA Ex: TA-23-0067)")

    directorate_options = [
        "ICT", "Finance & Accountant", "FMS GA", "Quality & Nursing",
        "Strategy & Commercial", "Medical", "Human Capital", "Legal",
        "Go beyond", "Management Secretary", "Corporate Strategy & Sustainability",
        "Network Operation", "Network Development", "Archetype", "SMC", "GKCI",
        "AMNT", "STC"
    ]
    directorate_choice = st.selectbox("Direktorat", directorate_options)
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Tanggal Mulai")
    with col2:
        end_date = st.date_input("Tanggal Selesai")

if st.button("🚀 Submit Semua Tanggal", use_container_width=True):
    if not name or not nik:
        st.warning("⚠️ Harap isi Nama dan NIK terlebih dahulu")
    else:
        weekdays = get_weekdays(start_date.strftime("%m/%d/%Y"), end_date.strftime("%m/%d/%Y"))
        
        if not weekdays:
            st.warning("❌ Tidak ada hari kerja dalam rentang tanggal yang dipilih")
        else:
            st.info(f"📅 Akan mengisi form untuk {len(weekdays)} hari kerja")
            with st.spinner("⏳ Sedang memproses, harap tunggu..."):
                automate_form(emp_status, nik, name, directorate_choice, weekdays)