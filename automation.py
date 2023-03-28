import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd
import os

sender_email = 'Enter the Email from which you want to send the mail'
sender_password = 'xyz'


# Enter the name of your HTML Template
with open('Template.html') as f:
    email_template = f.read()

# Read recipients from an Excel or CSV file
recipients_df = pd.read_excel('recipients.xlsx')  # Replace with the path to your Excel or CSV file

# Get attachment files interactively
attachments = []
while True:
    attachment_path = input('Enter path to attachment file (or "done" if finished): ')
    if attachment_path == 'done':
        break
    if not os.path.exists(attachment_path):
        print('Error: File not found.')
    else:
        attachments.append(attachment_path)

# Merge recipients into attachments into a single dataframe
data = pd.merge(recipients_df, pd.DataFrame({'attachment': attachments}), left_index=True, right_index=True)

for _, row in data.iterrows():
    recipient_name = row['name']
    recipient_email = row['email']
    attachment_path = row['attachment']

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = 'xyz'

    email_message = email_template.replace('{{name}}', recipient_name)
    msg.attach(MIMEText(email_message, 'html'))

    with open(attachment_path, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype=os.path.splitext(attachment_path)[1][1:])
        attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
        msg.attach(attachment)

    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)

print("Emails sent successfully")

