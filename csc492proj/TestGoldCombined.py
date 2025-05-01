import pandas as pd
import smtplib
from email.message import EmailMessage
import os
import glob
import shutil
from datetime import datetime
from typing import cast
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

os.chdir("C:\\Users\\KACronin001\\Documents\\Gold Project") #Make this the directory of where the code is
#print("Current Working Directory:", os.getcwd()) #Uncomment to check where it's looking

# Define file path
processed_folder = "C:\\Users\\KACronin001\\Documents\\Gold Project\\EmailSentXLSX"     # Make sure this folder exists

courseNames_df = pd.read_csv("coursenames.csv")

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

    # Email you are sending from
    fromaddr = "convcsc2025project@gmail.com"

    smtp.login(fromaddr, "omyu xsxt ryrh uxbp")

    # CC and BCC email
    cc = ["someemailyouwanttocc@gmail.com"]
    bcc = ["someemailyouwanttobcc@gmail.com"]

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
        courseCode = courseCode[:7] #This basically removes the other stuff we dont need since we only need the 3 letters and 3 numbers for the course code (No Y1 and stuff like that)

        courseNames_df.iloc[:, 0] = courseNames_df.iloc[:, 0].astype(str) #This is to help avoid errors (hopefully)
        course_row = courseNames_df[courseNames_df.iloc[:, 0].str.strip() == courseCode] #str.strip is here to help with data validation (so random spaces don't interfere with the checking (hopefully))

        if not course_row.empty:
            courseName = course_row.iloc[0, 1]
            print(f"Course Name Found: {courseName}")
        else:
            courseName = "Unknown Course Name"
            print(f"Course name not found for code: {courseCode}")

        #print(courseCode) #Uncomment to see the courseCode
        
        dfTest = pd.read_excel(excel_file, 0, header=header_row)
        #print(dfTest.columns.tolist()) #Uncomment to see the names of the columns
        dfTest_csv = dfTest.copy()
        dfTest_csv['Email Sent'] = ''
        dfTest_csv['Sent Timestamp'] = ''

        for i in range(len(dfTest)):
            ID = dfTest.loc[i, 'Jenzabar ID']
            FN = dfTest.loc[i, 'FIRST NAME']
            LN = dfTest.loc[i, 'LAST NAME']
            email = dfTest.loc[i, 'email address']

            if pd.notnull(email):
                message = MIMEMultipart()
                message["From"] = fromaddr
                message["To"] = str(email)
                message["Subject"] = "Information for your Converse University course"
                message["Cc"] = ", ".join(cc)  # Add CC header

                # Message to be sent
                body = f"""\
                Welcome to {courseName}, {courseCode}. Your professor will be Dr. {prof_last_name} ({email}).

                Your course is being taught through Converse's Canvas. The simplest way to get in to Converse's Canvas is to go to https://converse.instructure.com and log in there with your Converse email and password.

                If you've taken a course with Converse before, your Converse email address and password will normally be unchanged from what they were. If you have a new Converse account, you should have gotten your Converse email address and password in a separate email from Campus Technology.

                Once logged in, you should be taken to your Canvas dashboard. On that dashboard, you should see a tile with the name of your course.

                If you don't see the tile before the course starts, that is not (yet) a problem. Our Canvas courses are created unpublished, which means they're hidden from students.

                Please remember that you can always email me (peter.brown@converse.edu) for Canvas questions. Dr. {prof_last_name} is a better source for all other questions.

                Peace,
                —Peter Brown

                Peter H. Brown, Ph.D.
                Asst. Professor of Computer Science
                Director of Distance Education
                Converse University
                """
                message.attach(MIMEText(body, "plain"))

                # Sending the email
                all_recipients = [str(email)] + [str(addr) for addr in cc] + [str(addr) for addr in bcc]
                #smtp.sendmail(fromaddr, all_recipients, message.as_string()) #Uncomment to send emails

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
