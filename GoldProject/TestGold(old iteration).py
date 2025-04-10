import pandas as pd
import smtplib
from email.message import EmailMessage
#import os

#This is something I am trying to mess with to do directory stuff
#dir = os.listdir(path=r'C:\Users\KACronin001\Documents\Gold Project')
#print(dir) 

dfTest = pd.read_excel("Test1GoldProj.xlsx")

print(dfTest.columns)

ID = dfTest.loc[0, 'Jenzabar ID']
FN = dfTest.loc[0, 'FIRST NAME']
LN = dfTest.loc[0, 'LAST NAME']
email = dfTest.loc[0, 'email address']

# Creates SMTP session
smtp = smtplib.SMTP('smtp.gmail.com', 587)

# Start TLS for security
smtp.starttls()

# Authentication
smtp.login("convcsc2025project@gmail.com", "omyu xsxt ryrh uxbp")

# Message to be sent


# Sending the mail
i=0
for i in range(len(dfTest)):
  ID = dfTest.loc[i, 'Jenzabar ID']
  FN = dfTest.loc[i, 'FIRST NAME']
  LN = dfTest.loc[i, 'LAST NAME']
  email = dfTest.loc[i, 'email address']
  if pd.notnull(email):
    #The message stuff is so that it hopefully doesn't go to spam anymore.
    message = EmailMessage()
    message['From'] = "convcsc2025project@gmail.com"
    message['To'] = email
    message['Subject'] = "Test Email"
    message.set_content("I'm alive and coming for you.")
    smtp.send_message(message)
    # Print to show process has been completed
    print(f"Mail Sent To {email}")

# Terminating the session
smtp.quit()