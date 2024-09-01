import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import tkinter as tk
from tkinter import filedialog
from tkinter.simpledialog import askstring

# Function to open file explorer and select the Excel file
def get_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Function to get the custom message either from user input or a text file
def get_custom_message():
    choice = askstring("Input Method", "Type 'paste' to paste the message directly or 'file' to select a text file:")
    if choice.lower() == 'paste':
        return askstring("Custom Message", "Please paste your custom message:")
    elif choice.lower() == 'file':
        file_path = filedialog.askopenfilename(title="Select the Text File", filetypes=[("Text files", "*.txt")])
        with open(file_path, 'r') as file:
            return file.read()
    else:
        print("Invalid choice. Please restart the script and try again.")
        exit()

# Prompt user for the custom message
custom_message = get_custom_message()

# Prompt user for file location
print("Please select the Excel file containing the business names and phone numbers.")
file_path = get_file_path()

# Load the Excel file containing the business names and email addresses
df = pd.read_excel(file_path)

# Email setup
smtp_server = "https://mail.oracomgroup.com"
#smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "cedrick@oracomgroup.com"  # Replace with your email
password = "cedrick@34!"  # Replace with your email password

# Connect to the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Secure the connection
server.login(sender_email, password)

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    business_name = row['Business Name']
    recipient_email = row['Email 1']  # Assuming the primary email column is 'Email 1'
    
    # Personalize the message for each business
    personalized_message = f"Hello {business_name},\n{custom_message}"
    
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Your Subject Here"  # Replace with your subject
    msg.attach(MIMEText(personalized_message, 'plain'))

    # Send the email
    server.sendmail(sender_email, recipient_email, msg.as_string())

    # Wait for 25 seconds before sending the next email
    time.sleep(25)

# Close the server connection
server.quit()

print("Emails sent successfully.")
