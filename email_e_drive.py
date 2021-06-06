import os
import time
from pathlib import Path
from datetime import datetime
from collections import namedtuple
import pandas as pd

# import email, smtplib, ssl
# from email import encoders
# from email.mime.base import MIMEBase
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


File = namedtuple('File', 'name path size modified_date')

files = []

p = Path('C:')

# Gathers Data in the C: and outputs it to a CSV
for item in p.glob('**/*'):
    if item.suffix in (['.vib']):
        name = item.name
        path = Path.resolve(item.parent)
        size = item.stat().st_size
        modified = datetime.fromtimestamp(item.stat().st_mtime)

        files.append(File(name, path, size, modified))

df = pd.DataFrame(files)
df.to_csv('backups_c.csv', index=False)

#email function
def send_email(email_recipient, email_subject, email_message, attachment_location = ''):
    email_sender = 'youremail@yourdomain.com'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('youremail@yourdomain.com', 'YourPassword')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
    return True

#function call
send_email('youremail@yourdomain.com', 'Test subject', 'Test body', "C:/Python_Scripts/backups_c.csv")


#remove file after sending 
os.remove("C:/Python_Scripts/backups_c.csv")