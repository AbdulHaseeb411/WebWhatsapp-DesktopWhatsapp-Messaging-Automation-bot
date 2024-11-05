import sys
import time
import csv
from PyQt5 import QtWidgets
from pywinauto import Application, findwindows, keyboard
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class WhatsAppAutomationApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('WhatsApp Automation')
        self.setGeometry(100, 100, 400, 200)

        self.layout = QtWidgets.QVBoxLayout()

        self.label = QtWidgets.QLabel("Select WhatsApp Type:")
        self.layout.addWidget(self.label)

        self.radio_web = QtWidgets.QRadioButton("WhatsApp Web")
        self.radio_desktop = QtWidgets.QRadioButton("WhatsApp Desktop")
        self.layout.addWidget(self.radio_web)
        self.layout.addWidget(self.radio_desktop)

        self.csv_input = QtWidgets.QLineEdit(self)
        self.csv_input.setPlaceholderText("Enter CSV file path")
        self.layout.addWidget(self.csv_input)

        self.message_input = QtWidgets.QLineEdit(self)
        self.message_input.setPlaceholderText("Enter your message")
        self.layout.addWidget(self.message_input)

        self.send_button = QtWidgets.QPushButton("Send Messages")
        self.send_button.clicked.connect(self.send_messages)
        self.layout.addWidget(self.send_button)

        self.setLayout(self.layout)

    def send_messages(self):
        csv_filename = self.csv_input.text()
        message = self.message_input.text()
        if self.radio_web.isChecked():
            self.send_web_messages(csv_filename, message)
        elif self.radio_desktop.isChecked():
            self.send_desktop_messages(csv_filename, message)
        else:
            QtWidgets.QMessageBox.warning(self, "Warning", "Please select a WhatsApp type.")

    def send_web_messages(self, csv_filename, message):
        contacts = self.read_contacts_from_csv(csv_filename)
        driver = self.initialize_driver()
        driver.get('https://web.whatsapp.com/')
        time.sleep(50)  # Wait for QR code scan
        for contact_name in contacts:
            self.send_whatsapp_message(driver, contact_name, message)
        driver.quit()

    def send_desktop_messages(self, csv_filename, message):
        windows = findwindows.find_windows(title='WhatsApp', backend='uia')
        if windows:
            app = Application(backend='uia').connect(handle=windows[0])
            contacts = self.read_contacts_from_csv(csv_filename)
            for contact_name in contacts:
                self.send_desktop_message(app, contact_name, message)
        else:
            QtWidgets.QMessageBox.warning(self, "Warning", "WhatsApp Desktop is not running.")

    def send_whatsapp_message(self, driver, contact_name, message):
        search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
        search_box.click()
        search_box.clear()
        search_box.send_keys(contact_name)
        time.sleep(2)
        search_box.send_keys(Keys.ENTER)
        time.sleep(5)
        message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true" and @data-tab="10"]')
        message_box.click()
        message_box.send_keys(message)
        time.sleep(3)
        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(send_button)).click()
        time.sleep(3)

    def send_desktop_message(self, app, contact_name, message):
        search_box = app.window(title='WhatsApp').descendants(control_type="Edit")[0]
        search_box.set_focus()
        search_box.type_keys(contact_name, with_spaces=True)
        time.sleep(3)
        keyboard.send_keys("{ENTER}")
        time.sleep(5)

        # Check the available edit controls
        message_input_controls = app.window(title='WhatsApp').descendants(control_type="Edit")
        print(f"Found {len(message_input_controls)} edit controls.")

        if len(message_input_controls) > 1:
            message_input_box = message_input_controls[1]
            message_input_box.set_focus()
            message_input_box.type_keys(message, with_spaces=True)
            keyboard.send_keys("{ENTER}")
            time.sleep(3)
            keyboard.send_keys("{ESC}")
        else:
            print("Message input box not found. Make sure WhatsApp is open and active.")

    def read_contacts_from_csv(self, filename):
        contacts = []
        with open(filename, newline='') as csvfile:
            csv_reader = csv.reader(csvfile)
            for row in csv_reader:
                contacts.append(row[0].strip())
        return contacts

    def initialize_driver(self):
        chrome_options = Options()
        chrome_options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        return driver

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = WhatsAppAutomationApp()
    window.show()
    sys.exit(app.exec_())
