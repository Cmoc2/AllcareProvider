import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import shutil
import re
import tkinter
from tkinter import filedialog
import sys


def send_email_with_files(source_folder, email_mapping, sender_email, app_password, sent_folder):
    # Ensure the sent folder exists
    if not os.path.exists(sent_folder):
        os.makedirs(sent_folder)
    
    # Iterate over all files in the source folder
    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)
        if os.path.isfile(file_path):
            # Extract the patient's name from the filename
            match = re.match(r"^(.*?)(\d{8})", filename)
            if match:
                patient_name = match.group(1).strip()
            else:
                patient_name = "Unknown"

            # Determine the recipient email based on the filename
            recipient_email = None
            if "(ECAH)" in filename:
                recipient_email = 'kp.ecah@allcareprovider.com'
            else:
                for keyword, email in email_mapping.items():
                    if keyword in filename:
                        recipient_email = email
                        break
            
            if recipient_email:
                # Determine the subject based on the filename
                if "ROC Order" in filename:
                    subject = f'ROC Order: {patient_name}'
                elif "Order -" in filename:
                    subject = f'Order: {patient_name}'
                else:
                    subject = f'File: {patient_name}'

                # Create the email message
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = subject

                # Attach the file
                attachment = MIMEBase('application', 'octet-stream')
                with open(file_path, 'rb') as file:
                    attachment.set_payload(file.read())
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', f'attachment; filename={filename}')
                msg.attach(attachment)

                # Connect to the SMTP server and send the email
                try:
                    with smtplib.SMTP('smtp.office365.com', 587) as server:
                        server.starttls()
                        server.login(sender_email, app_password)
                        server.sendmail(sender_email, recipient_email, msg.as_string())
                        print(f'Email sent successfully to {recipient_email} for file {filename}')
                        
                        # Move the file to the sent folder
                        shutil.move(file_path, os.path.join(sent_folder, filename))
                        print(f'Moved {filename} to {sent_folder}')
                except Exception as e:
                    print(f'Failed to send email for file {filename}. Reason: {e}')
            else:
                print(f'No matching email found for file {filename}')

# Example usage
source_folder = 'C:/Users/ChristianOrtiz/Documents/OrdersToEmail/Pending'
sent_folder = 'C:/Users/ChristianOrtiz/Documents/OrdersToEmail/Done'
email_mapping = {
    'Order - BP': 'baldwin.park@allcareprovider.com',
    'Order - LA': 'la.metro@allcareprovider.com',
    'Order - D': 'downey@allcareprovider.com',
    'Order - SB': 'south.bay@allcareprovider.com',
    'Order - V': 'valley@allcareprovider.com',
    'Order - F': 'fontana@allcareprovider.com',
    'Order - OC': 'orange.county@allcareprovider.com',
    'Order - CL': 'carelon@allcareprovider.com',
    'Order - Test': 'christian.ortiz@allcareprovider.com'

}
sender_email = 'christian.ortiz@allcareprovider.com'
app_password = 'lqqctxcnkmhvpbwp'

# Display Message to select a Folder
tkinter.messagebox.showinfo(title="Folder Select", message="Select the folder containing Orders to email.")

# Show Dialog box and return path of folder
source_folder = filedialog.askdirectory(title='Select a folder')

# Check if folder was selected
if source_folder:
    print(f"Selected folder: {source_folder}")
    # Display Message to select output Folder
    tkinter.messagebox.showinfo(title="Folder Select", message="Select the folder for successfully sent orders.")

    # Show Dialog box and return path of folder
    sent_folder = filedialog.askdirectory(title='Select a folder')
    send_email_with_files(source_folder, email_mapping, sender_email, app_password, sent_folder)
    tkinter.messagebox.showinfo(title="Operation Complete", message="Orders Sent.")
else:
    print("No Folder Selected.")
    # Display No Folder Selected Message
    tkinter.messagebox.showinfo(title="No Input/Output Folder Selected", message="Operation Aborted.")
    sys.exit(0)
    
