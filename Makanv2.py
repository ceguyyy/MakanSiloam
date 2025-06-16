import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import time

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

def automate_form(emp_status, nik, name, directorate_choice, dates):
    FORM_URL = "https://forms.office.com/pages/responsepage.aspx?id=UwNJH5TnM0e5oqJO5GV_T9KMRjpWHsZEkKVvBch9KKlUOFVBNFA1MEdMU0xKVFFZVjAzV0tPV1ZDSi4u&route=shorturl"

    status_xpath_map = {
        "Employee": '//*[@id="question-list"]/div[1]/div[2]/div/div/div[1]/div/label/span[1]/input',
        "Vendor/Outsource": '//*[@id="question-list"]/div[1]/div[2]/div/div/div[2]/div/label/span[1]/input',
        "Non Employee": '//*[@id="question-list"]/div[1]/div[2]/div/div/div[3]/div/label/span[1]/input'
    }

    for date_str in dates:
        st.write(f"Processing for {date_str}...")

        driver = webdriver.Chrome()
        driver.get(FORM_URL)
        time.sleep(3)

        try:
            # Select Employment Status
            driver.find_element(By.XPATH, status_xpath_map[emp_status]).click()

            # Input NIK
            driver.find_element(By.XPATH, '/html/body/div/div/div[1]/div/div/div/div/div[3]/div/div/div[2]/div[2]/div[2]/div[2]/div/span/input').send_keys(nik)

            # Input Name
            driver.find_element(By.XPATH, '//*[@id="question-list"]/div[3]/div[2]/div/span/input').send_keys(name)

            # Select Directorate
            dir_xpath = f'//*[@id="question-list"]/div[4]/div[2]/div/div/div[{directorate_choice}]/div/label/span[1]/input'
            driver.find_element(By.XPATH, dir_xpath).click()

            # Fill Date
            date_input = driver.find_element(By.XPATH, '//*[@id="DatePicker0-label"]')
            date_input.send_keys(date_str)
            date_input.send_keys(Keys.TAB)

            time.sleep(2)

            # Submit
            submit_btn = driver.find_element(By.XPATH, '//*[@id="form-main-content1"]/div/div/div[2]/div[3]/div/button')
            submit_btn.click()

            st.success(f"Submitted form for {date_str}")

        except Exception as e:
            st.error(f"Error on {date_str}: {e}")
        finally:
            time.sleep(2)
            driver.quit()

# === STREAMLIT UI ===
st.title("Mempermudah Makan Anda di SHOO")

emp_status = st.selectbox("Select Employment Status", ["Employee", "Vendor/Outsource", "Non Employee"])
name = st.text_input("Enter your full name")
nik = st.text_input("Enter NIK (for TA = ex : TA-23-0067)")

directorate_options = [
    "ICT", "Finance & Accountant", "FMS GA", "Quality & Nursing",
    "Strategy & Commercial", "Medical", "Human Capital", "Legal",
    "Go beyond", "Management Secretary", "Corporate Strategy & Sustainability",
    "Network Operation", "Network Development", "Archetype", "SMC", "GKCI",
    "AMNT", "STC"
]
directorate_choice = st.selectbox("Select Directorate", directorate_options)
start_date = st.date_input("Start Date")
end_date = st.date_input("End Date")

if st.button("Submit All Dates"):
    if not name or not nik:
        st.warning("Please fill in both Name and NIK.")
    else:
        dir_index = directorate_options.index(directorate_choice) + 1
        weekdays = get_weekdays(start_date.strftime("%m/%d/%Y"), end_date.strftime("%m/%d/%Y"))
        automate_form(emp_status, nik, name, dir_index, weekdays)
