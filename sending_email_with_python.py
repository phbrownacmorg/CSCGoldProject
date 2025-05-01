import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# Creates SMTP session
smtp = smtplib.SMTP('smtp.gmail.com', 587)

# Start TLS for security
smtp.starttls()

# Email you are sending from
fromaddr = "convcsc2025project@gmail.com"

# Email you are sending to
toaddr = "someemailyouwanttosendto@gmail.com"

# Authentication
smtp.login(fromaddr, "omyu xsxt ryrh uxbp")

# CC and BCC email
cc = ["someemailyouwanttocc@gmail.com"]
bcc = ["someemailyouwanttobcc@gmail.com"]

course_title = "placeholdercoursetitle"
prof_last_name = "placeholderproflastname"
instructor_email = "placeholder@gmail.com"
message = "Information for your Converse University course"

message = MIMEMultipart()
message["From"] = fromaddr
message["To"] = toaddr
message["Cc"] = ", ".join(cc) # Add CC header
message["Bcc"] = ", ".join(bcc) # Add BCC header
message["Subject"] = "Information for your Converse University course"

# Message to be sent
body = f"""\
Welcome to {section}, {course_title}. Your professor will be Dr. {prof_last_name} ({instructor_email}).

Your course is being taught through Converse's Canvas. The simplest way to get in to Converse's Canvas is to go to https://converse.instructure.com and log in there with your Converse email and password. (Please note that you have to use your Converse email address for this. A personal email, or a school-district email, won't work.)

If you've taken a course with Converse before, your Converse email address and password will normally be unchanged from what they were. If you have a new Converse account, you should have gotten your Converse email address and password in a separate email from Campus Technology. If, for any reason, you don't have your address and password, you can get your Converse email and password by contacting Campus Technology at 864-596-9457 during their business hours (M-Th 8-5, F 8-1) or at helpdesk@converse.edu.

Once logged in, you should be taken to your Canvas dashboard. On that dashboard, you should see a tile with the name of your course in the middle of the page. Click that tile to be taken to the course.

[This paragraph only appears before the course's starting date.]
If you don't see the tile before the course starts, that is not (yet) a problem. Our Canvas courses are created unpublished, which means they're hidden from students. The professor publishes each course when it's ready. So if you don't see your course immediately, that merely means that your professor hasn't published it yet. If you're still not seeing your course by a day after it's supposed to start, please contact your professor. If that doesn't help, contact me (Dr. Peter Brown, peter.brown@converse.edu).

Please remember that you can always email me (peter.brown@converse.edu) for Canvas questions. Dr. {prof_last_name} is a better source for all other questions.

Peace,
—Peter Brown

Peter H. Brown, Ph.D.
Asst. Professor of Computer Science
Director of Distance Education
Converse University
"""

message.attach(MIMEText(body, "plain"))

# Sending the mail
all_recipients = [toaddr] + cc + bcc
smtp.sendmail(fromaddr, all_recipients, message.as_string())
#smtp.sendmail("convcsc2025project@gmail.com", toaddrs, message)

# Print to show process has been completed
print("Mail Sent")

# Terminating the session
smtp.quit()
