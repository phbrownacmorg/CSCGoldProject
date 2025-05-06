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

def workingDir():
    working_dir: str = os.path.join(os.path.expanduser('~'), 'Documents', "Gold Project")
    if not os.path.isdir(working_dir):
        working_dir = os.path.dirname(__file__)
    print('Working dir: ', working_dir)
    os.chdir(working_dir) #Make this the directory of where the code is
    #print("Current Working Directory:", os.getcwd()) #Uncomment to check where it's looking
    return working_dir

def folders(working_dir):
    # Define file path
    input_folder = os.path.join(working_dir, "EmailListingInput")
    assert os.path.isdir(input_folder), str(input_folder) + ' does not exist'  # Make sure this folder exists

    processed_folder = os.path.join(working_dir, "EmailListingProcessed")
    os.makedirs(processed_folder,exist_ok=True)

    csvFolder = os.path.join(working_dir, "EmailSentCSV") #Name this whatever the name of the folder is
    os.makedirs(csvFolder,exist_ok=True)

    return input_folder,processed_folder,csvFolder

def csvData(working_dir):
    courseNames_csv = os.path.join(working_dir, "coursenames.csv")
    print(courseNames_csv)
    assert os.path.isfile(courseNames_csv), courseNames_csv + ' does not exist'  # Make sure this file exists
    courseNames_df = pd.read_csv(courseNames_csv)

    instructors_csv = os.path.join(working_dir, "instructors.csv")
    print(instructors_csv)
    assert os.path.isfile(instructors_csv), instructors_csv + ' does not exist'  # Make sure this file exists
    instructors_df = pd.read_csv(instructors_csv)

    return courseNames_df,instructors_df

def smtpSetup():
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()

    # Email you are sending from
    fromaddr = "convcsc2025project@gmail.com"
    smtp.login(fromaddr, "omyu xsxt ryrh uxbp")

    return smtp, fromaddr

def sendEmail(smtp, fromaddr, cc_email, bcc_emails, courseName, courseCode, prof_last_name, emails):
    message = MIMEMultipart()
    message["From"] = fromaddr
    message["To"] = ", ".join(emails)
    message["Subject"] = "Information for your Converse University course"
    message["Cc"] = ", ".join(cc_email)  # Add CC header

    # Message to be sent
    body = f"""\
    Welcome to {courseName}, {courseCode}. Your professor will be Dr. {prof_last_name} ({cc_email}).

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
    all_recipients = emails + cc_email + bcc_emails
    #smtp.sendmail(fromaddr, all_recipients, message.as_string()) #Uncomment to send emails

def excelStuff(input_folder, processed_folder, csvFolder, smtp, fromaddr, courseNames_df, instructors_df):
    excel_files = glob.glob('*.xlsx', root_dir=input_folder) + glob.glob('*.xls', root_dir=input_folder)

    if not excel_files:
        print("No Excel files found in {0}.".format(input_folder))
        return
    else:
        print(f"Found {len(excel_files)} Excel file(s):")

        for excel_file in excel_files:
            print(f"Processing: {excel_file}")
        
            base_name = os.path.splitext(excel_file)[0]

            parts = base_name.split('_')
            
            if parts and ',' in parts[-1]:
                prof_last_name = parts[-1].split(',')[0]
            else:
                prof_last_name = 'Unknown'

            print(f"Professor Last Name Detected: {prof_last_name}")
            instructor_row = instructors_df[instructors_df['lastname'] == prof_last_name]
            if not instructor_row.empty:
                prof_email = instructor_row.iloc[0]['email']
                print(f"Professor Email Found: {prof_email}")
                cc_email = [prof_email]  # Assign this email to CC
            else:
                print(f"No email found for Professor {prof_last_name}. No CC will be added.")
                cc_email = []
            
            #This code block before dfTest is so that it can find where the actual header row is as the files don't have them in the same area
            temp = pd.read_excel(os.path.join(input_folder, excel_file), header=None, nrows=5)

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
            dfTest = pd.read_excel(os.path.join(input_folder, excel_file), 0, header=header_row)
            #print(dfTest.columns.tolist()) #Uncomment to see the names of the columns
            dfTest_csv = dfTest.copy()
            dfTest_csv['Email Sent'] = ''
            dfTest_csv['Sent Timestamp'] = ''

            #Realized the email should be only 1 email instead of sending a bunch of emails to different students
            #dropna() removes NaN/missing values
            #unique() makes sure there are no duplicate emails (I think I remember one of our examples before having a duplicate)
            #toList() makes it a list
            emails = dfTest['email address'].dropna().unique().tolist()
            bcc_emails = dfTest['School EMAIL'].dropna().unique().tolist() 
            
            sendEmail(smtp, fromaddr, cc_email, bcc_emails, courseName, courseCode, prof_last_name, emails)

            for i in range(len(dfTest)):
                email = dfTest.loc[i, 'email address']
                if pd.notnull(email):
                    dfTest_csv.at[i, 'Email Sent'] = 'Sent'
                    dfTest_csv.at[i, 'Sent Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    print(f"Mail Sent To {email}")
            
            csv_path = os.path.join(csvFolder, base_name + '.csv')
            dfTest_csv.to_csv(csv_path, index=False)
            print(f"Saved CSV to: {csv_path}")
            shutil.move(os.path.join(input_folder, excel_file), os.path.join(processed_folder, os.path.basename(excel_file)))
            print(f"Moved {excel_file} to {processed_folder}")

def main():
    working_dir = workingDir()
    courseNames_df, instructors_df = csvData(working_dir)
    input_folder, processed_folder, csvFolder = folders(working_dir)
    smtp, fromaddr = smtpSetup()
    cc = ["someemailyouwanttocc@gmail.com"] 

    excelStuff(input_folder, processed_folder, csvFolder,smtp, fromaddr, courseNames_df, instructors_df)

    smtp.quit()

if __name__ == "__main__":
    main()