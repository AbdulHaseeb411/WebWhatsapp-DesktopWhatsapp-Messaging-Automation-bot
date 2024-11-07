import sys
import time
import csv
from PyQt5 import QtWidgets, QtGui, QtCore
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
        # Window properties
        self.setWindowTitle("WhatsApp Automation")
        self.setGeometry(100, 100, 500, 700)
        self.setStyleSheet("background-color: #e0cab8")

        # Layout
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.setAlignment(QtCore.Qt.AlignCenter)

        # Title Label
        title_label = QtWidgets.QLabel("WhatsApp Automation Tool", self)
        title_label.setStyleSheet("""
            color: #4CAF50;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 20px;
        """)
        title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.layout.addWidget(title_label)

        # Radio Buttons for WhatsApp Type
        self.radio_web = QtWidgets.QRadioButton("WhatsApp Web", self)
        self.radio_desktop = QtWidgets.QRadioButton("WhatsApp Desktop", self)
        self.radio_web.setStyleSheet("color: #4CAF50; font-size: 18px; font-weight: bold;")
        self.radio_desktop.setStyleSheet("color: #4CAF50; font-size: 18px; font-weight: bold;")
        self.layout.addWidget(self.radio_web)
        self.layout.addWidget(self.radio_desktop)

        # Input Fields and Send Button
        self.csv_input = QtWidgets.QLineEdit(self)
        self.csv_input.setPlaceholderText("Enter or browse CSV file path")
        self.csv_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #4CAF50;
                border-radius: 8px;
                font-size: 18px;
                font-weight: bold;
                background-color: #ffffff;
                color: #4CAF50;
            }
            QLineEdit:hover {
                border-color: #3e8e41;
            }
        """)
        self.layout.addWidget(self.csv_input)

        # Browse button
        self.browse_button = QtWidgets.QPushButton("Browse", self)
        self.browse_button.setStyleSheet("""
            QPushButton {
                padding: 8px;
                font-size: 18px;
                font-weight: bold;
                background-color: #4CAF50;
                color: #ffffff;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #3e8e41;
            }
        """)
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)

        # Message input
        self.message_input = QtWidgets.QLineEdit(self)
        self.message_input.setPlaceholderText("Enter your message")
        self.message_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #4CAF50;
                border-radius: 8px;
                font-size: 18px;
                font-weight: bold;
                background-color: #ffffff;
                color: #4CAF50;
            }
            QLineEdit:hover {
                border-color: #1976D2;
            }
        """)
        self.layout.addWidget(self.message_input)

        # Send Button
        self.send_button = QtWidgets.QPushButton("Send Messages")
        self.send_button.setStyleSheet("""
            QPushButton {
                background-color: #7DF9FF;
                color: #ffffff;
                font-size: 18px;
                font-weight: bold;
                padding: 10px;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #00FFFF;
            }
        """)
        self.send_button.clicked.connect(self.send_messages)
        self.layout.addWidget(self.send_button)

    def send_messages(self):
        try:
            csv_filename = self.csv_input.text()
            message = self.message_input.text()

            if not csv_filename or not message:
                raise ValueError("CSV file path and message cannot be empty.")

            if self.radio_web.isChecked():
                self.send_web_messages(csv_filename, message)
            elif self.radio_desktop.isChecked():
                self.send_desktop_messages(csv_filename, message)
            else:
                raise ValueError("Please select a WhatsApp type.")
        except Exception as e:
            self.show_error(str(e))

    def send_web_messages(self, csv_filename, message):
        try:
            contacts = self.read_contacts_from_csv(csv_filename)
            driver = self.initialize_driver()
            driver.get('https://web.whatsapp.com/')
            time.sleep(50)  # Wait for QR code scan
            for contact_name in contacts:
                self.send_whatsapp_message(driver, contact_name, message)
            driver.quit()
        except Exception as e:
            self.show_error(f"Error sending web messages: {str(e)}")

    def send_desktop_messages(self, csv_filename, message):
        try:
            # Attempt to find the WhatsApp window
            windows = findwindows.find_windows(title='WhatsApp', backend='uia')

            if windows:
                # Connect to the WhatsApp window and send messages
                app = Application(backend='uia').connect(handle=windows[0])
                contacts = self.read_contacts_from_csv(csv_filename)
                for contact_name in contacts:
                    self.send_desktop_message(app, contact_name, message)
            else:
                self.show_error("WhatsApp Desktop is not running.")
        except Exception as e:
            self.show_error(f"Error sending desktop messages: {str(e)}")

    def browse_file(self):
        """Open file dialog to select a CSV file."""
        try:
            file_dialog = QtWidgets.QFileDialog(self)
            file_dialog.setFileMode(QtWidgets.QFileDialog.ExistingFiles)
            file_dialog.setNameFilter("CSV files (*.csv)")
            file_dialog.setViewMode(QtWidgets.QFileDialog.List)

            if file_dialog.exec_():
                selected_files = file_dialog.selectedFiles()
                if selected_files:
                    self.csv_input.setText(selected_files[0])
        except Exception as e:
            self.show_error(f"Error browsing file: {str(e)}")

    def show_error(self, message):
        QtWidgets.QMessageBox.critical(self, "Error", message)

    def read_contacts_from_csv(self, filename):
        contacts = []
        try:
            with open(filename, newline='') as csvfile:
                csv_reader = csv.reader(csvfile)
                for row in csv_reader:
                    contacts.append(row[0].strip())
            return contacts
        except Exception as e:
            self.show_error(f"Error reading CSV file: {str(e)}")
            return []

    def send_whatsapp_message(self, driver, contact_name, message):
        try:
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
        except Exception as e:
            self.show_error(f"Error sending message to {contact_name}: {str(e)}")

    def send_desktop_message(self, app, contact_name, message):
        try:
            search_box = app.window(title='WhatsApp').descendants(control_type="Edit")[0]
            search_box.set_focus()
            search_box.type_keys(contact_name, with_spaces=True)
            time.sleep(5)
            keyboard.send_keys("{ENTER}")
            time.sleep(5)

            matching_chat_item = self.find_matching_chat_item(app, contact_name)
            if matching_chat_item:
                time.sleep(2)
            else:
                return

            message_input_controls = app.window(title='WhatsApp').descendants(control_type="Edit")
            for control in message_input_controls:
                if control.is_visible() and control.is_enabled():
                    message_input_box = control
                    break
            if message_input_box:
                message_input_box.set_focus()
                message_input_box.type_keys(message, with_spaces=True)
                keyboard.send_keys("{ENTER}")
                time.sleep(3)
            else:
                return
        except Exception as e:
            self.show_error(f"Error sending desktop message to {contact_name}: {str(e)}")

    def find_matching_chat_item(self, app, contact_name):
        try:
            main_window = app.window(title='WhatsApp')
            time.sleep(2)
            chat_items = main_window.descendants(control_type='ListItem')
            for item in chat_items:
                if contact_name.lower() in item.window_text().lower():
                    item.click_input(double=True)
                    time.sleep(2)
                    return True
            return False
        except Exception as e:
            self.show_error(f"Error finding chat item: {str(e)}")
            return False

    def initialize_driver(self):
        try:
            options = Options()
            options.add_argument('--user-data-dir=C:\\path\\to\\chrome\\user\\data')
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            return driver
        except Exception as e:
            self.show_error(f"Error initializing WebDriver: {str(e)}")
            return None

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = WhatsAppAutomationApp()
    ex.show()
    sys.exit(app.exec_())
