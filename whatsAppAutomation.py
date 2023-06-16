import logging
from selenium import webdriver
import datetime
import re
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials

logging.basicConfig(level=logging.INFO)

# Set up Google Sheets API credentials
scope = ['https://www.googleapis.com/auth/spreadsheets']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/...credentials.json', scope)
client = gspread.authorize(credentials)

# Open the Google Sheets workbook
spreadsheet = client.open_by_key('....')

# Select the appropriate sheet within the workbook
worksheet = spreadsheet.worksheet('[name of the sheet]')

survey_link = "[link]"
message_lines = [
    " ",
    " ",
    " ",
    " ",
    " ",
    " ",
    survey_link
]

# Set up Chrome profile directory
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=/Users/.../Library/Application Support/Google/Chrome/Default")

# Set the path to the ChromeDriver executable
chromedriver_path = '/Users/.../chromedriver'
webdriver_service = Service(chromedriver_path)

# Pass the options and service to the ChromeDriver
driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)


# Function to check if 30 days have passed since the start date
def is_past_30_days(date_string):
    if date_string == '?' or date_string == '':
        return False

    try:
        current_date = datetime.datetime.now().date()
        date = datetime.datetime.strptime(date_string, "%d/%m/%Y").date()
        days_passed = (current_date - date).days
        return days_passed >= 30
    except ValueError:
        return False


# Function to send WhatsApp message
def send_whatsapp_message(phone_number):
    logging.info(f"Sending WhatsApp message to phone number: {phone_number}")
    # Set up WhatsApp message content
    try:
        # Wait for the QR code scan
        WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'p.selectable-text.copyable-text.iq0m558w')))
        logging.info("WhatsApp Web loaded successfully")

        # Search for the phone number in the chat list
        search_box = driver.find_element(By.CSS_SELECTOR, 'div[data-tab="3"]')
        search_box.clear()
        search_box.send_keys(phone_number)

        # Find the chat element
        chat_xpath = '//div[@class="_13jwn"]'
        chat_element = WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, chat_xpath)))

        # Click on the chat element
        chat_element.click()

        # Find the chat input box
        chat_input_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'div[data-testid="conversation-compose-box-input"][role="textbox"]')))

        # Clear the input box (optional)
        chat_input_box.clear()

        # Click on the input box to focus on it
        chat_input_box.click()

        # Send each line of the message
        for line in message_lines:
            chat_input_box.send_keys(line)
            chat_input_box.send_keys(Keys.SHIFT, Keys.ENTER)  # Press Shift+Enter for a line break

        # Submit the message by pressing the "Enter" key
        chat_input_box.send_keys(Keys.ENTER)

        logging.info(f"Message sent to phone number: {phone_number}")

        # Clear the search box
        search_box.clear()

    except TimeoutException:
        logging.error("Failed to load WhatsApp Web.")
        return False


# Main script
logging.info("Fetching the latest data from the Google Sheet")
try:
    driver.get('https://web.whatsapp.com')
    logging.info("Opened WhatsApp Web")

    worksheet = spreadsheet.worksheet("...")
    data = worksheet.get_all_values()

    for row in data[8:]:
        logging.info(f"Processing row: {row}")

        name = row[1]  # Update column index to 1 (column B)
        start_date = row[2]  # Update column index to 2 (column C)
        message_sent = row[10]  # Update column index to 10 (column K)
        phone_number = row[9]  # Update column index to 9 (column J)

        # Check if the name is empty
        if name == "":
            logging.info("Name is empty. Skipping to the next row.")
            continue

        # Check if the phone number is empty or doesn't match the pattern
        if phone_number == "" or not re.match(r"\+\d{3} \d{2} \d{3} \d{4}", phone_number):
            logging.info(f"Invalid phone number: {phone_number}. Skipping to the next row.")
            continue

        # Check if the message has already been sent
        if message_sent == "נשלח":
            logging.info(f"Message already sent for phone number: {phone_number}")
            continue

        # Check if the start date is valid and 30 days have passed since the start date
        if start_date != "" and is_past_30_days(start_date) and message_sent != "נשלח":
            logging.info(f"Sending WhatsApp message for phone number: {phone_number}")
            send_whatsapp_message(phone_number)  # Pass the phone_number as an argument

            # Check if the message was sent successfully
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'p.selectable-text.copyable-text.iq0m558w')))
            logging.info("WhatsApp message sent successfully")

            # Find the cell based on the name
            cell = worksheet.find(name)
            worksheet.update_cell(cell.row, 11, "נשלח")  # Update column index to 11 (column K)
            logging.info(f"Updated column K for phone number: {phone_number} to 'נשלח'")

    print("Sending WhatsApp From Google Sheet Api Script Done.")

    # Close the WebDriverWait
    driver.quit()

except Exception as e:
    logging.error(f"An error occurred: {str(e)}")
