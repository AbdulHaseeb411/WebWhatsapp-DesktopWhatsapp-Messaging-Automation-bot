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
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
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

    def send_desktop_messages(self, csv_filename, message):
        windows = findwindows.find_windows(title="WhatsApp", backend="uia")
        if not windows:
            messagebox.showerror("Error", "WhatsApp Desktop is not running.")
            return

        app = Application(backend="uia").connect(handle=windows[0])
        contacts = self.read_contacts_from_csv(csv_filename)
        
        for contact_name in contacts:
            self.send_desktop_message(app, contact_name, message)

    def read_contacts_from_csv(self, filename):
        contacts = []
        with open(filename, newline="") as csvfile:
            csv_reader = csv.reader(csvfile)
            for row in csv_reader:
                contacts.append(row[0].strip())
        return contacts

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
        search_box = app.window(title="WhatsApp").descendants(control_type="Edit")[0]
        search_box.set_focus()
        search_box.type_keys(contact_name, with_spaces=True)
        time.sleep(5)
        
        keyboard.send_keys("{ENTER}")
        time.sleep(5)
        
        message_input_box = app.window(title="WhatsApp").descendants(control_type="Edit")[-1]
        message_input_box.set_focus()
        message_input_box.type_keys(message, with_spaces=True)
        keyboard.send_keys("{ENTER}")
        time.sleep(3)

    def initialize_driver(self):
        options = Options()
        options.add_argument("--user-data-dir=C:\\path\\to\\chrome\\user\\data")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        return driver

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAutomationApp(root)
    root.mainloop()
