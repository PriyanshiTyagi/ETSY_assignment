import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Loading the Excel sheet
wb = openpyxl.load_workbook('seller_data.xlsx')
sheet = wb.active

# Setting up the Selenium webdriver
driver = webdriver.Chrome()

# Iterate through the rows in the sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    seller_name, store_link = row

    # Visiting the store
    driver.get(store_link)

    try:
        # Finding the "message" button and click it
        message_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Message')]"))
        )
        message_button.click()

        # Finding the message input field and send the template message
        message_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//textarea[@name='message']"))
        )
        message_input.send_keys("Hello, I have some query about your products.")

        # Find the send button and click it
        send_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Send')]"))
        )
        send_button.click()

    except Exception as e:
        print(f"An error occurred while sending a message to {seller_name}: {str(e)}")

# Close the browser
driver.quit()
