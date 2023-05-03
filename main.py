import os
import smtplib
import imaplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta


# self defined method. It could delete certain "format" files whose created time less than "days" and under "path"
# folder.
def delete_old_files(path, days, format):
    # get current date
    now = datetime.now()

    # go through all files under path
    for filename in os.listdir(path):
        # check is the files' format correct
        if filename.endswith("." + format):
            file_path = os.path.join(path, filename)

            # get the files' created dates
            file_creation_time = os.path.getctime(file_path)
            file_creation_date = datetime.fromtimestamp(file_creation_time)

            # delete files outdated
            if (now - file_creation_date) > timedelta(days=days):
                os.remove(file_path)
                print(f"Deleted: {file_path}")


# configure for IMAP to get mails from inbox
imap_server = "your.imap.server"
smtp_server = "your.smtp.server"
username = "your@email.com"
password = "your_password"
recipient = "recipient@email.com"

# connect to IMAP server
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(username, password)

# choose inbox
mail.select("INBOX")

# search for mails only for "today"
today = datetime.today().strftime("%d-%b-%Y")
result, data = mail.search(None, f'SINCE "{today}"')

# go through all target mails
email_ids = data[0].split()
for num in email_ids:
    result, msg_data = mail.fetch(num, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])

    # download attachments files
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if not filename or not filename.endswith(".xls"):
            continue

        # for long-term use, I'd like to store the xls files in local directory and delete outdated files(for
        # instance, I would delete txt and xls files out of 7 days). I defined a method at the end of the page let's
        # xsassume the path is “local\Download”
        path = "local\Download"
        delete_old_files(path, 7, "txt")
        delete_old_files(path, 7, "xls")

        # delete the outdated files, then download today's new xls files
        filepath = os.path.join("local\Download", filename)
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))

        # a fake "operation" here. If we do need to operate or edit xls, I would use pandas and marplot libs.
        txt_filename = filename.replace(".xls", ".txt")
        filepath = os.path.join("attachments", txt_filename)
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))

# create new mails and add "operated" files as attachments
msg = MIMEMultipart()
msg["From"] = username
msg["To"] = recipient
msg["Subject"] = "Converted TXT Files"
body = "Please find the attached converted txt files."
msg.attach(MIMEText(body, "plain"))

# only accepts files whose created time less than 1 day. Add them in msg.attach()
for filename in os.listdir("attachments"):
    if filename.endswith(".txt"):
        file_path = os.path.join("attachments", filename)
        file_time = datetime.fromtimestamp(os.path.getctime(file_path))
        if (datetime.now() - file_time).days <= 1:
            with open(file_path, "r") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={filename}")
                msg.attach(part)

# send email
server = smtplib.SMTP_SSL(smtp_server)
server.login(username, password)
server.sendmail(username, recipient, msg.as_string())
server.quit()
