import time
import csv
from pywinauto import Application, findwindows, keyboard

def find_search_box(app):
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

def find_matching_chat_item(app, contact_name):
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

def find_message_input_box(app):
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

def read_contacts_from_csv(file_path):
    """Read contacts from a CSV file."""
    contacts = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if row:  # Avoid empty rows
                contacts.append(row[0].strip())  # Add the contact name
    return contacts

def go_back_to_chat_list(app):
    """Navigate back to the main chat list in WhatsApp."""
    try:
        main_window = app.window(title='WhatsApp')
        main_window.set_focus()
        keyboard.send_keys("{ESC}")  # Use ESC to go back to the chat list
        time.sleep(2)  # Wait for the chat list to be ready
        print("Navigated back to chat list.")
    except Exception as e:
        print(f"Could not navigate back to chat list: {e}")

def main():
    try:
        windows = findwindows.find_windows(title='WhatsApp', backend='uia')
        if windows:
            app = Application(backend='uia').connect(handle=windows[0])
            print("WhatsApp Desktop opened successfully.")
            
            # Read contacts from the CSV file
            contacts = read_contacts_from_csv('contect.csv')
            print(f"Loaded contacts: {contacts}")

            # Prompt for the message to send
            message_to_send = input("Enter the message you want to send: ")

            for contact_name in contacts:
                search_box = find_search_box(app)
                if search_box:
                    search_box.set_focus()
                    search_box.type_keys(contact_name, with_spaces=True)
                    time.sleep(3)  # Allow time for search results to load

                    keyboard.send_keys("{ENTER}")
                    print(f"Opened chat with {contact_name}.")

                    matching_chat_item = find_matching_chat_item(app, contact_name)
                    if matching_chat_item:
                        print(f"Clicked on the chat item for {contact_name}.")
                        time.sleep(2)  # Wait for chat to open
                            
                        # Find the message input box and send the message
                        message_input_box = find_message_input_box(app)
                        if message_input_box:
                            message_input_box.set_focus()  # Set focus on the message input box
                            message_input_box.type_keys(message_to_send, with_spaces=True)  # Type the message
                            keyboard.send_keys("{ENTER}")  # Send the message
                            print(f"Message sent to {contact_name}: {message_to_send}")
                        else:
                            print("Message input box not found.")
                    else:
                        print(f"No chat item to click for {contact_name}.")
                
                # Go back to the main chat list before the next contact
                go_back_to_chat_list(app)

                # Optional: Add a delay between messages
                time.sleep(3)

        else:
            print("WhatsApp Desktop is not running.")
            return

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
