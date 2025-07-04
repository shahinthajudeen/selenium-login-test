from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from datetime import datetime
import time
import webbrowser
import os

# Setup ChromeDriver path
driver_path = r"C:\Windows\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Track login status
login_success = False

try:
    driver.get("https://www.saucedemo.com/")
    driver.maximize_window()

    # Wait for elements and perform login
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "user-name"))
    ).send_keys("standard_user")

    driver.find_element(By.ID, "password").send_keys("secret_sauce")
    driver.find_element(By.ID, "login-button").click()

    # Wait for redirect to inventory page
    WebDriverWait(driver, 10).until(
        EC.url_contains("inventory")
    )

    # Confirm login by checking URL and page content
    if "inventory" in driver.current_url and "inventory" in driver.page_source:
        login_success = True

        # Show success popup on page
        driver.execute_script("""
            var msg = document.createElement('div');
            msg.innerText = ' Login Successful';
            msg.style.position = 'fixed';
            msg.style.top = '20px';
            msg.style.right = '20px';
            msg.style.backgroundColor = '#28a745';
            msg.style.color = 'white';
            msg.style.padding = '12px 20px';
            msg.style.borderRadius = '8px';
            msg.style.fontSize = '18px';
            msg.style.fontWeight = 'bold';
            msg.style.boxShadow = '0 4px 6px rgba(0,0,0,0.1)';
            msg.style.zIndex = '9999';
            document.body.appendChild(msg);
        """)

except Exception as e:
    print(f"Error occurred: {e}")

finally:
    # Prepare timestamps
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    safe_time = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Create HTML report
    html_content = f"""
    <html>
    <head><title>Selenium Test Report</title></head>
    <body style='font-family:Arial;'>
        <h2>SauceDemo Login Test Report</h2>
        <p><strong>Date:</strong> {current_time}</p>
        <p><strong>Status:</strong> {"Passed" if login_success else "Failed"}</p>
        <p><strong>Page Title:</strong> {driver.title}</p>
        <p><strong>URL:</strong> {driver.current_url}</p>
    </body>
    </html>
    """

    with open("login_test_report.html", "w") as f:
        f.write(html_content)

    webbrowser.open("login_test_report.html")

    # Create Excel report
    wb = Workbook()
    ws = wb.active
    ws.title = "Login Test Result"
    ws.append(["Date", "Test Name", "Status", "Title", "URL"])

    status = "Passed" if login_success else "Failed"
    ws.append([current_time, "SauceDemo Login Test", status, driver.title, driver.current_url])

    # Save Excel to Desktop
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    excel_filename = f"login_test_result_{safe_time}.xlsx"
    full_excel_path = os.path.join(desktop_path, excel_filename)
    wb.save(full_excel_path)

    os.startfile(full_excel_path)

    print(login_success)
    input("Press Enter to quit...")
    driver.quit()
