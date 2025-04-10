import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import shutil
from datetime import datetime


# File paths
input_folder = r"C:\Users\rogle\OneDrive\Documents\Class stuff\GoldProject\EmailListingInput"
processed_folder = r"C:\Users\rogle\OneDrive\Documents\Class stuff\GoldProject\EmailListingProcessed"
filename = "Test1GoldProj.xlsx"
input_filepath = os.path.join(input_folder, filename)
processed_filepath = os.path.join(processed_folder, filename)

# Timestamp for new filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
processed_filename = f"Test1GoldProj.xlsx"
processed_filepath = os.path.join(processed_folder, processed_filename)

# Read Excel file
dfTest = pd.read_excel(input_filepath)
print(dfTest.columns)

# Add missing columns if they don't exist
if 'email_sent' not in dfTest.columns:
    dfTest['email_sent'] = False
if 'sent_date' not in dfTest.columns:
    dfTest['sent_date'] = ""

# Email setup
smtp = smtplib.SMTP('smtp.gmail.com', 587)
smtp.starttls()
smtp.login("convcsc2025project@gmail.com", "omyu xsxt ryrh uxbp")

# Send emails and update columns
for i in range(len(dfTest)):
    email = dfTest.loc[i, 'email address']
    already_sent = dfTest.loc[i, 'email_sent']

    # Only send if email is valid and not already sent
    if pd.notnull(email) and not already_sent:
        try:
            # Construct email
            message = EmailMessage()
            message['From'] = "convcsc2025project@gmail.com"
            message['To'] = email
            message['Subject'] = "Test Email"
            message.set_content("I'd like to apologize for the previous rick roll. I am testing the input to processed stuff. You will be rick rolled several more times though.")

            # Send the message
            smtp.send_message(message)
            print(f"Mail Sent To {email}")

            # Update status
            dfTest.at[i, 'email_sent'] = True
            dfTest.at[i, 'sent_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        except Exception as e:
            print(f"Failed to send to {email}: {e}")

# Close email session
smtp.quit()

# Save updated Excel as a new file
os.makedirs(processed_folder, exist_ok=True)
dfTest.to_excel(processed_filepath, index=False)
print(f"Updated file saved as {processed_filepath}")