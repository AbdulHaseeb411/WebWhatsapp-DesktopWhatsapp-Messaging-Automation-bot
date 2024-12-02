import tkinter as tk
from tkinter import filedialog, messagebox, Tk, Frame, Label, Entry, Button, StringVar, ttk
import csv
import time
from pywinauto import Application, findwindows, keyboard
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image, ImageTk 
import re
import subprocess
from pywinauto import Desktop
import xml.etree.ElementTree as ET
class WhatsAppAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Automation Tool")
        self.root.geometry("1000x700")

        # Load and set the background image
        self.background_image = Image.open("image.png")  # Replace with your image file
        self.background_image = self.background_image.resize((800, 700), Image.LANCZOS)
        self.bg_image = ImageTk.PhotoImage(self.background_image)

        # Create a label to hold the background image and place it
        self.bg_label = Label(root, image=self.bg_image)
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Title
        title_label = Label(root, text="WhatsApp Automation Tool", font=("Helvetica", 12, "bold"), fg="#333333", bg="#f4f4f9")
        title_label.pack(pady=20)  # Center alignment with padding

        # Dropdown for WhatsApp type with Label in a vertical column
        platform_label = Label(root, text="Select WhatsApp  Type:", font=("Helvetica", 10, "bold"), fg="#555555", bg="#f4f4f9")
        platform_label.pack(anchor="w", padx=20)

        self.platform_var = StringVar()
        self.platform_dropdown = ttk.Combobox(root, textvariable=self.platform_var, font=("Helvetica", 10, "bold"), state="readonly", width=20)
        self.platform_dropdown["values"] = ["WhatsApp Web", "WhatsApp Desktop"]
        self.platform_dropdown.pack(anchor="w", padx=20, pady=5)

        # File Input and Browse Button in a vertical column
        csv_label = Label(root, text="CSV File Path:", font=("Helvetica", 12, "bold"), fg="#555555", bg="#f4f4f9")
        csv_label.pack(anchor="w", padx=20, pady=(10, 0))

        self.csv_input = Entry(root, font=("Helvetica", 10), width=30, fg="#333333", bg="#ffffff")
        self.csv_input.insert(0, "Enter or browse CSV file path")
        self.csv_input.pack(anchor="w", padx=20, pady=(0, 5))

        self.browse_button = Button(root, text="Browse", font=("Helvetica", 10, "bold"), fg="#ffffff", bg="#4CAF50", command=self.browse_file)
        self.browse_button.pack(anchor="w", padx=20, pady=(0, 15))

        # Message Input with Label in a vertical column
        message_label = Label(root, text="Enter Message:", font=("Helvetica", 12), fg="#555555", bg="#f4f4f9")
        message_label.pack(anchor="w", padx=20, pady=(10, 0))

        self.message_input = Entry(root, font=("Helvetica", 10), width=30, fg="#333333", bg="#ffffff")
        self.message_input.pack(anchor="w", padx=20, pady=(0, 15))

        # Send Button
        self.send_button = Button(root, text="Send Messages", font=("Helvetica", 10, "bold"), fg="#ffffff", bg="#3e8e41", command=self.send_messages)
        self.send_button.pack(anchor="w", padx=20, pady=20)




    def browse_file(self):
        # Allow selection of both CSV and XML files
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("XML files", "*.xml")])
        
        if file_path:
            # Clear the current input field and insert the selected file path
            self.csv_input.delete(0, tk.END)
            self.csv_input.insert(0, file_path)


    def send_messages(self):
        csv_filename = self.csv_input.get()
        message = self.message_input.get()
        platform = self.platform_var.get()

        if not csv_filename or not message or not platform:
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            if platform == "WhatsApp Web":
                self.send_web_messages(csv_filename, message)
            elif platform == "WhatsApp Desktop":
                self.send_desktop_messages(csv_filename, message)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def send_web_messages(self, csv_filename, message):
        contacts = self.read_contacts_from_csv(csv_filename)
        driver = self.initialize_driver()
        driver.get("https://web.whatsapp.com/")
        time.sleep(50)
        
        for contact_name in contacts:
            self.send_whatsapp_message(driver, contact_name, message)
        driver.quit()

    def open_whatsapp_desktop(self):
   
        try:
            # Access the taskbar and search for WhatsApp
            taskbar = Desktop(backend="uia").window(class_name="Shell_TrayWnd")
            whatsapp_button = taskbar.child_window(title="WhatsApp", control_type="Button")
    
            # Click on the WhatsApp taskbar icon
            whatsapp_button.click_input()
            
            # Allow time for WhatsApp to open
            time.sleep(5)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open WhatsApp from the taskbar: {e}")
            return False
    def send_desktop_messages(self, csv_filename, message):
        # Check if WhatsApp is running
        windows = findwindows.find_windows(title="WhatsApp", backend="uia")
        if not windows:
            if not self.open_whatsapp_desktop():
                return

        # Recheck after attempting to open
        windows = findwindows.find_windows(title="WhatsApp", backend="uia")
        if not windows:
            messagebox.showerror("Error", "Unable to open or detect WhatsApp Desktop.")
            return

        # Connect to WhatsApp Desktop
        app = Application(backend="uia").connect(handle=windows[0])
        contacts = self.read_contacts_from_csv(csv_filename)
        
        for contact_name in contacts:
            self.send_desktop_message(app, contact_name, message)

    def read_contacts(self, filename):
        contacts = []

        # Check file extension to determine the format (CSV or XML)
        if filename.endswith(".csv"):
            contacts = self.read_contacts_from_csv(filename)
        elif filename.endswith(".xml"):
            contacts = self.read_contacts_from_xml(filename)
        else:
            messagebox.showerror("Error", "Unsupported file format. Please use CSV or XML.")

        return contacts

    def read_contacts_from_csv(self, filename):
        contacts = []
        try:
            with open(filename, newline="") as csvfile:
                csv_reader = csv.reader(csvfile)
                for row in csv_reader:
                    contacts.append(row[0].strip())  # Assuming contacts are in the first column
        except FileNotFoundError:
            messagebox.showerror("Error", f"CSV file '{filename}' not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading the CSV file: {e}")

        return contacts

    def read_contacts_from_xml(self, filename):
        contacts = []
        try:
            tree = ET.parse(filename)
            root = tree.getroot()

            # Assuming contacts are stored as 'contact' elements under the root
            for contact in root.findall("contact"):
                name = contact.text.strip()  # Get the contact name
                contacts.append(name)
        except FileNotFoundError:
            messagebox.showerror("Error", f"XML file '{filename}' not found.")
        except ET.ParseError:
            messagebox.showerror("Error", f"Error parsing the XML file '{filename}'.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading the XML file: {e}")

        return contacts

    import re

    def send_whatsapp_message(self, driver, contact_name, message):
        time.sleep(4)

        # Locate the search box and input the contact name
        search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
        search_box.click()
        search_box.clear()
        search_box.send_keys(contact_name)
        time.sleep(2)

        try:
            # Locate all chat results matching the search
            search_results = driver.find_elements(By.XPATH, '//span[contains(@title, "")]')

            if not search_results:
                print(f"No search results found for '{contact_name}'. Skipping to the next contact.")
                return

            # Normalize contact_name to avoid hidden characters
            contact_name = re.sub(r'\s+', ' ', contact_name.strip().lower())  # Remove extra spaces and lower case

            # Iterate through all search results to find the exact title match
            for result in search_results:
                chat_name = result.get_attribute("title").strip().lower()  # Normalize chat_name

                # Debugging log to see what chat_name is being checked
                print(f"Checking chat name: {chat_name}")

                # Check if the title matches the contact name exactly
                if re.sub(r'\s+', ' ', chat_name) == contact_name:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", result)  # Ensure visibility
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(result)).click()  # Click
                        time.sleep(5)  # Wait for chat to open

                        # Locate the message input box and send the message
                        message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true" and @data-tab="10"]')
                        message_box.click()
                        message_box.send_keys(message)
                        time.sleep(3)

                        # Locate and click the send button
                        send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send"]')
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(send_button)).click()
                        print(f"Message sent to {contact_name}.")
                        time.sleep(3)
                        return  # Exit after sending the message to the correct contact
                    except Exception as click_error:
                        print(f"Error while clicking on '{chat_name}': {str(click_error)}")
                        continue  # Try the next result if clicking fails

            # If no exact match is found, log and skip
            print(f"Contact '{contact_name}' not found in exact search results. Skipping to the next contact.")

        except Exception as e:
            print(f"Error occurred while searching for contact '{contact_name}': {str(e)}")


    def find_message_input_box(self, app):
            """Helper function to locate the message input box."""
            message_input_controls = app.window(title="WhatsApp").descendants(control_type="Edit")
            print(f"Found {len(message_input_controls)} edit controls.")

            if len(message_input_controls) > 1:
                return message_input_controls[-1]  # Assuming the last one is the message input box
            return None
    def send_desktop_message(self, app, contact_name, message):
        main_window = app.window(title="WhatsApp")
        time.sleep(2)  # Allow time for WhatsApp to load

        # Step 1: Find the search box
        print(f"Locating the search box...")
        search_box = self.find_search_box(app)
        if search_box:
            print("Search box found, proceeding with search.")
            # Step 2: Search for the contact
            print(f"Searching for contact '{contact_name}'...")
            search_box.set_focus()
            search_box.type_keys(contact_name, with_spaces=True)
            time.sleep(5)  # Wait for search results to load

            # Step 3: Find the matching chat item
            matching_chat_item = self.find_matching_chat_item(app, contact_name)

            if matching_chat_item:
                print(f"Clicked on the chat item.")
                time.sleep(2)  # Allow time for the chat to open

                # Step 4: Locate the message input box and send the message
                message_input_box = self.find_message_input_box(app)
                if message_input_box and message_input_box.is_enabled() and message_input_box.is_visible():
                    message_input_box.set_focus()
                    message_input_box.type_keys(message, with_spaces=True)
                    keyboard.send_keys("{ENTER}")
                    print(f"Message sent.")
                    keyboard.send_keys("{ESC}")
                    
                else:
                    print("Message input box not found or not interactable.")
            else:
                print(f"No chat found for '{contact_name}'.")
        else:
            print("Search box not found.")

    def find_search_box(self, app):
        """Try to locate the search box in WhatsApp."""
        try:
            main_window = app.window(title='WhatsApp')
            edit_controls = main_window.descendants(control_type="Edit")
            print(f"Found {len(edit_controls)} Edit controls.")

            for i, control in enumerate(edit_controls):
                print(f"Trying Edit control index {i}")
                control.set_focus()
                if control.is_visible():
                    print("Search box found successfully.")
                    return control

            print("No suitable search box found.")
            return None

        except Exception as e:
            print(f"Could not find the search box: {e}")
            return None

    def find_matching_chat_item(self, app, contact_name):
        """Find the first chat item that matches the contact name."""
        try:
            main_window = app.window(title='WhatsApp')
            time.sleep(2)  # Wait for chat list to populate
            chat_items = main_window.descendants(control_type='ListItem')
            print(f"Found {len(chat_items)} chat items.")

            for chat in chat_items:
                chat_text = chat.window_text().strip()  # Get the text of the chat item
                print(f"Checking chat item: {chat_text}")  # Print the text for debugging

                if contact_name.lower() in chat_text.lower():  # Check if the contact name is in the chat item
                    print(f"Found matching chat item: {chat_text}")
                    chat.click_input()  # Click on the matching chat item
                    return chat  # Return the matching chat item

            print("No matching chat item found.")
            return None

        except Exception as e:
            print(f"Could not find chat items: {e}")
            return None

    def find_message_input_box(self, app):
        """Find the message input box in WhatsApp."""
        try:
            main_window = app.window(title='WhatsApp')
            time.sleep(2)  # Wait for message input box to become visible
            edit_controls = main_window.descendants(control_type="Edit")
            for control in edit_controls:
                control.set_focus()
                if control.is_visible() and control.is_enabled():
                    print("Message input box found.")
                    return control  # Return the first visible and enabled Edit control as the message input box

            print("No suitable message input box found.")
            return None

        except Exception as e:
            print(f"Could not find message input box: {e}")
            return None







    def initialize_driver(self):
        options = Options()
        options.add_argument("--user-data-dir=C:\\path\\to\\chrome\\user\\data")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        return driver

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAutomationApp(root)
    root.mainloop()
