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
toaddr = "linklamoreaux2002@gmail.com"

# Authentication
smtp.login(fromaddr, "omyu xsxt ryrh uxbp")

# CC and BCC email
cc = ["dinosaurchainsaw2002@gmail.com"]
bcc = ["ljlamoreaux001@converse.edu"]

# Message to be sent
message = MIMEMultipart()
message["From"] = fromaddr
message["To"] = toaddr
message["Cc"] = ", ".join(cc) # Add CC header
message["Bcc"] = ", ".join(bcc) # Add BCC header
message["Subject"] = "Test Email"

body = "This is the email body."
message.attach(MIMEText(body, "plain"))

# Sending the mail
all_recipients = [toaddr] + cc + bcc
smtp.sendmail(fromaddr, all_recipients, message.as_string())
#smtp.sendmail("convcsc2025project@gmail.com", toaddrs, message)

# Print to show process has been completed
print("Mail Sent")

# Terminating the session
smtp.quit()
