import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

load_dotenv()
df = pd.read_excel('companies.xlsx')

sender_email = os.getenv('SENDER_EMAIL')
password = os.getenv('PASSWORD')


def send_email(to_email, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype=os.path.splitext(attachment_path)[1][1:])
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
            msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587) 
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.close()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {e}")

# Function to read email template from file
def read_template(file_path):
    with open(file_path, 'r') as file:
        template = file.read()
    return template

# Read email template from file
template_file = 'email_template.md'  # Change to the actual file name and path
email_template = read_template(template_file)

for index, row in df.iterrows():
    company_name = row['Company Name']
    to_email = row['Email']
    job_vacancy = row['Job Vacancy']
    
    resume_file = 'resume.pdf' 
    
    subject = f"Application for {job_vacancy} at {company_name}"
    body = email_template.format(company_name=company_name, job_vacancy=job_vacancy)

    send_email(to_email, subject, body, resume_file)

# Made By Anoob Suresh