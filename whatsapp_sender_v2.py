import pandas as pd
import time
from datetime import datetime
import keyboard
import os
import sys
import webbrowser
import pyperclip
import urllib.parse
import pyautogui

def read_message_template(template_file):
    """Read the message template from file"""
    with open(template_file, 'r', encoding='utf-8') as file:
        return file.read()

def format_phone_number(phone, name):
    """Format phone number and add Singapore country code"""
    phone = ''.join(filter(str.isdigit, phone))
    
    # Handle different phone number formats
    if phone.startswith('65'):
        phone = '+' + phone
    elif phone.startswith('+65'):
        pass  # Already in correct format
    else:
        phone = '+65' + phone
    
    # Validate phone number length for Singapore (8 digits + country code)
    if len(phone) != 11:  # +65 + 8 digits
        raise ValueError(f"Invalid phone number length for {name}: {phone}")
    
    return phone

def send_whatsapp_messages(excel_file, template_file):
    """
    Send personalized WhatsApp messages to contacts using a single WhatsApp Web tab.
    """
    try:
        # Read the message template
        message_template = read_message_template(template_file)
        
        # Read the Excel file
        df = pd.read_excel(excel_file, names=['Phone', 'Name'])
        
        print(f"\nFound {len(df)} contacts in the Excel file.")
        print("\nInstructions:")
        print("1. When WhatsApp Web opens, scan the QR code")
        print("2. Wait for WhatsApp Web to fully load")
        print("3. Press Enter in this window to start sending messages")
        input("\nPress Enter when ready to start...")
        
        # Open WhatsApp Web in default browser
        #webbrowser.open('https://web.whatsapp.com')
        #time.sleep(5)  # Wait for browser to open
        
        # Wait for user to scan QR code and load WhatsApp Web
        #input("\nAfter scanning QR code and WhatsApp Web is loaded, press Enter to continue...")
        
        # Process each contact
        for index, row in df.iterrows():
            try:
                # Get phone number and name
                phone = str(row['Phone']).strip()
                name = str(row['Name']).strip()
                
                # Format phone number
                phone = format_phone_number(phone, name)
                
                # Personalize message
                personalized_message = message_template.replace('<Name>', name)
                
                print(f"\nPreparing message for {name} ({phone})")
                
                # Create WhatsApp URL for this contact
                message_encoded = urllib.parse.quote(personalized_message)
                whatsapp_url = f'https://web.whatsapp.com/send?phone={phone}&text={message_encoded}'
                
                # Copy message to clipboard (as backup in case URL encoding fails)
                pyperclip.copy(personalized_message)
                
                # Open chat with this contact
                webbrowser.open(whatsapp_url)
                
                print("Waiting for chat to load...")
                time.sleep(10)  # Wait for chat to load
                
                # Press Enter to send message
                keyboard.press_and_release('enter')
                
                print(f"✓ Message sent to {name}")
                
                # Wait before next message
                time.sleep(5)
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(2)
                
            except Exception as e:
                print(f"❌ Error sending message to {name} ({phone}): {str(e)}")
                continue
            
    except Exception as e:
        print(f"Error processing files: {str(e)}")
        input("\nPress Enter to exit...")
        sys.exit(1)

def main():
    print("WhatsApp Bulk Message Sender")
    print("===========================")
    print("\nMake sure you have:")
    print("1. contact_list.xlsx file (with Phone and Name columns)")
    print("2. message.txt file (with your message template)")
    print("\nBoth files should be in the same folder as this program.")
    
    # Check if files exist
    if not os.path.exists("contact_list.xlsx"):
        print("\n❌ Error: contact_list.xlsx not found!")
        input("Press Enter to exit...")
        sys.exit(1)
        
    if not os.path.exists("message.txt"):
        print("\n❌ Error: message.txt not found!")
        input("Press Enter to exit...")
        sys.exit(1)
    
    print("\n✓ Found required files")
    input("Press Enter to continue...")
    
    send_whatsapp_messages("contact_list.xlsx", "message.txt")
    
    print("\nAll messages have been processed!")
    input("Press Enter to exit...")

if __name__ == "__main__":
    # Install required packages:
    # pip install pandas pyperclip keyboard openpyxl
    main()