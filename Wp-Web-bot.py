import time
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def initialize_driver():
    """Initialize the Chrome driver with specified options."""
    chrome_options = Options()
    chrome_options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # Adjust if necessary
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def send_whatsapp_message(driver, contact_name, message):
    """Send a WhatsApp message to the specified contact."""
    # Search for the contact or group
    search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
    search_box.click()
    search_box.clear()
    search_box.send_keys(contact_name)
    time.sleep(2)  # Wait for search results to load

    # Press Enter to select the contact
    search_box.send_keys(Keys.ENTER)
    print(f"Opened chat with {contact_name}.")

    # Wait for the chat to load
    time.sleep(5)  # Adjust wait time as necessary

    # Locate the message box
    message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true" and @data-tab="10"]')
    message_box.click()
    message_box.send_keys(message)
    
    # Wait a moment before sending the message
    time.sleep(2)

    # Locate the send button
    send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')

    # Wait until the send button is clickable and try clicking it
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send"]')))
        send_button.click()  # Click the send button to send the message
        print(f"Message sent to {contact_name}: {message}")
        
        # Additional wait to ensure the message is sent
        time.sleep(3)  # Wait to allow the message to send
    except Exception as click_error:
        print(f"Error sending message to {contact_name}: {click_error}")

def read_contacts_from_csv(filename):
    """Read contact names from a CSV file."""
    contacts = []
    with open(filename, newline='') as csvfile:
        csv_reader = csv.reader(csvfile)
        for row in csv_reader:
            contacts.append(row[0])  # Assuming contact names are in the first column
    return contacts

def main():
    """Main function to run the WhatsApp messaging script."""
    csv_filename = input("Enter the name of the CSV file with contacts: ")  # Get CSV file name from user
    message = input("Enter your message: ")  # Get message from user

    # Read contacts from CSV
    contacts = read_contacts_from_csv(csv_filename)
    print(f"Loaded {len(contacts)} contacts from {csv_filename}.")

    driver = None  # Define driver initially as None

    try:
        driver = initialize_driver()
        driver.get('https://web.whatsapp.com/')
        print("Chrome opened successfully. Please scan the QR code to log in to WhatsApp.")
        
        # Wait for the user to scan the QR code
        time.sleep(50)  # Adjust based on how long it takes to scan

        # Loop through each contact and send the message
        for contact_name in contacts:
            try:
                send_whatsapp_message(driver, contact_name, message)
                time.sleep(5)  # Small delay between messages
            except Exception as e:
                print(f"Failed to send message to {contact_name}: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close driver only if it was successfully created
        if driver:
            driver.quit()
            print("Driver closed.")
        else:
            print("Driver was not initialized.")

if __name__ == "__main__":
    main()
