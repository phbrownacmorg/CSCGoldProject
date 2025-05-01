import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import glob
import shutil
from datetime import datetime
from typing import cast

os.chdir("C:\\Users\\KACronin001\\Documents\\Gold Project") #Make this the directory of where the code is
#print("Current Working Directory:", os.getcwd()) #Uncomment to check where it's looking

# Define file path
processed_folder = "C:\\Users\\KACronin001\\Documents\\Gold Project\\EmailSentXLSX"     # Make sure this folder exists

csvFolder = "EmailSentCSV" #Name this whatever the name of the folder is
os.makedirs(csvFolder,exist_ok=True)

instructors_df = pd.read_csv("instructors.csv")

excel_files = glob.glob('*.xlsx') + glob.glob('*.xls')

if not excel_files:
    print("No Excel files found in the current directory.")
else:
    print(f"Found {len(excel_files)} Excel file(s):")

    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()
    smtp.login("convcsc2025project@gmail.com", "omyu xsxt ryrh uxbp")

    for excel_file in excel_files:
        print(f"Processing: {excel_file}")
        
        base_name = os.path.splitext(excel_file)[0]

        parts = base_name.split('_')
        
        if parts and ',' in parts[-1]:
            prof_last_name = parts[-1].split(',')[0]
        else:
            prof_last_name = 'Unknown'

        print(f"Professor Last Name Detected: {prof_last_name}")

        temp = pd.read_excel(excel_file, header=None, nrows=5) #This code block before dfTest is so that it can find where the actual header row is as the files don't have them in the same area

        header_row: int = -1
        for i, row in temp.iterrows():
            if 'Jenzabar ID' in row.values:
                header_row = cast(int, i)
                break

        if header_row == -1:
            print(f"Header not found in {excel_file}.")
            continue

        courseCode = temp.iloc[0,0]
        courseCode = str(courseCode) 
        courseCode = courseCode.replace(" ","") #This removes spaces
        courseCode = courseCode[:6] #This basically removes the other stuff we dont need since we only need the 3 letters and 3 numbers for the course code (No Y1 and stuff like that)
        #print(courseCode) #Uncomment to see the name of the courseCode
        dfTest = pd.read_excel(excel_file, 0, header=header_row)
        #print(dfTest.columns.tolist()) #Uncomment to see the names of the columns
        dfTest_csv = dfTest.copy()
        dfTest_csv['Email Sent'] = ''
        
        for i in range(len(dfTest)):
            ID = dfTest.loc[i, 'Jenzabar ID']
            FN = dfTest.loc[i, 'FIRST NAME']
            LN = dfTest.loc[i, 'LAST NAME']
            email = dfTest.loc[i, 'email address']
            if pd.notnull(email):
                message = EmailMessage()
                message['From'] = "convcsc2025project@gmail.com"
                message['To'] = email
                message['Subject'] = "Test Email"
                message.set_content("Hi, this is a test email.")
                #smtp.send_message(message) #Uncomment to actually send the emails.
                dfTest_csv.at[i, 'Email Sent'] = 'Sent'
                dfTest_csv.at[i, 'Sent Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"Mail Sent To {email}")

        csv_path = os.path.join(csvFolder, base_name + '.csv')
        dfTest_csv.to_csv(csv_path, index=False)
        print(f"Saved CSV to: {csv_path}")

        #Name this whatever the name of the folder is and put the path to the folder of origin
        excelFolder = os.path.join("C:\\Users\\KACronin001\\Documents\\Gold Project", "EmailSentXLSX")
        os.makedirs(excelFolder, exist_ok=True)

        #Move the excel file from the main folder to the EmailSentXLSX folder
        shutil.move(excel_file, os.path.join(excelFolder, os.path.basename(excel_file)))
        print(f"Moved {excel_file} to {excelFolder}")

    smtp.quit()
