import win32com.client
import logging

# Logging configuration
logging.basicConfig(
    filename='log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

try: 
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox folder
    
    logging.info("Connection to Outlook established.")
    for item in inbox.Items:
        if item.Class == 43:  # 43 = MailItem
            print(f"From: {item.SenderName}, Subject: {item.Subject}, Received: {item.ReceivedTime}")
            logging.info(f"Email found: From: {item.SenderName}, Subject: {item.Subject}, Received: {item.ReceivedTime}")
        elif item.Class == 26:  # 26 = AppointmentItem
            print(f"Appointment: {item.Subject}, Starts: {item.Start}, Ends: {item.End}")
            logging.info(f"Appointment found: Subject: {item.Subject}, Starts: {item.Start}, Ends: {item.End}")
        else:
            print("Item is not an email.")
            logging.info("Item ignored, not an email.")
except Exception as e:
    logging.error(f"Error connecting to Outlook: {e}")
    logging.exception(e)
    print(f"Error connecting to Outlook: {e}")

# Test connection to Outlook and list emails
print("Outlook connection test completed.")
print(f"Total emails in inbox: {len(inbox.Items)}")
print("End of test script.")
