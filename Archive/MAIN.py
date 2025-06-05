import logging
import pandas as pd
import time
from queue import Queue
from threading import Thread
import win32com.client as client
import pythoncom
import os
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pytz



# Setup logging
logging.basicConfig(filename='email_sending.log', level=logging.INFO)

logging.info("lets go")



# Load configuration
config = {
    'excel_path': '',
    'word_document_path': '',
    'html_path': 'Email_short_3.html',
    'email_account': 'jsb@synsol.ch',
    'subject_line': 'New Competitive Urea Solutions: Discover Our Indonesia to Europe Route',
    'send_time': '2024-02-02 07:00', # 'YYYY-MM-DD HH:MM' ZEITVERSCHIEBUNG hat Einfluss!!!! Winter 1h früher Sommer 1h später
    'sheet_name' : 'test'
    # ... other configurations
}

# Email queue
email_queue = Queue()

# Function to get all file paths in the "attachments" folder
def get_attachment_paths():
    attachments_dir = os.path.join(os.getcwd(), 'attachments')  # Get the path of the "attachments" folder
    attachment_paths = [os.path.join(attachments_dir, f) for f in os.listdir(attachments_dir) if os.path.isfile(os.path.join(attachments_dir, f))]
    return attachment_paths



def personalize_content(file_path, email_info):
    try:
        # Open and read the HTML file directly with BeautifulSoup
        with open(file_path, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')

        # For each placeholder, find and replace in the HTML content
        placeholders = {
            '[Vorname]': email_info.get('vorname', ''),
            '[Nachname]': email_info.get('nachname', ''),
            '[Title]': email_info.get('title', ''),
            '[Company]': email_info.get('company', ''),
        }
        
        for placeholder in list(placeholders):
            value = placeholders[placeholder]
            # Convert pandas NaN to an empty string
            if pd.isna(value):
                value = ''
            else:
                value = str(value)
            placeholders[placeholder] = value

            text_elements = soup.find_all(text=lambda text: placeholder in text)
            for element in text_elements:
                updated_text = element.replace(placeholder, value)
                element.replace_with(updated_text)

        return str(soup)
    except Exception as e:
        logging.error(f"Error personalizing HTML content: {e}")
        return ""


# Modify the send_email_from_queue function
def send_email_from_queue():
    # Initialize the COM library for the new thread
    pythoncom.CoInitialize()
    outlook = client.Dispatch('Outlook.Application')

    # Select the desired account
    account_to_use = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == config['email_account']:  # Replace with the email address of the account you want to use
            account_to_use = account
            break

    if account_to_use is None:
        logging.error("Failed to find the specified account.")
        return  # Exit the function if the specified account is not found
    
    # Parse the initial send time from the config
    send_time = datetime.strptime(config['send_time'], '%Y-%m-%d %H:%M')
    send_time = pytz.timezone('Europe/Zurich').localize(send_time)  # Adjust 'Your/Timezone' to your actual timezone, e.g., 'Europe/Zurich'
    

    email_index = 0  # To track the number of emails processed

    while True:
        email_info = email_queue.get()
        try:
            # Personalize the email content directly from the HTML file
            personalized_html_content = personalize_content(config['html_path'], email_info)
            # Create the email message
            if personalized_html_content:
                message = outlook.CreateItem(0)
                message.HTMLBody = personalized_html_content  # Set the HTML content as the body
                message.Subject = email_info['email_subject']
                message.To = email_info['email']
                # ... (potentially set other email fields like CC, BCC, etc.)
                # Set the account to use for sending
                if account_to_use is not None:
                    message.SendUsingAccount = account_to_use
                
                # Attach files from the "attachments" folder
                attachment_paths = get_attachment_paths()  # Get all attachment file paths
                for attachment_path in attachment_paths:
                    message.Attachments.Add(attachment_path)
                
                # Calculate and set the deferred delivery time
                current_email_send_time = send_time + timedelta(seconds=2.5 * email_index)
                message.DeferredDeliveryTime = current_email_send_time

                # Send the email
                message.Send()
                logging.info(f"Email sent to {email_info['email']}")
            else:
                logging.error(f"Failed to convert Word document to HTML for {email_info['email']}")
            
        except Exception as e:
            logging.error(f"Failed to send email to {email_info['email']}: {e}")
        finally:
            email_queue.task_done()
            email_index += 1
            time.sleep(2.5)  # Wait for 2.5 seconds before sending the next email


# Start the email sending thread
Thread(target=send_email_from_queue, daemon=True).start()

# Load emails from Excel, specifying the sheet name
leads_df = pd.read_excel(config['excel_path'], sheet_name=config['sheet_name'])

# Enqueue emails
for index, row in leads_df.iterrows():
    email_info = {
        'email': row['E-Mail'],
        'vorname': row['Vorname'],
        'nachname': row['Nachname'],
        'title': row['Title'],
        'company': row['Company'],
        'email_subject': config['subject_line'],  
        # ... other personalization fields

    }
    email_queue.put(email_info)

# Wait for the email queue to be processed
email_queue.join()
logging.info("All emails have been processed. new 2")