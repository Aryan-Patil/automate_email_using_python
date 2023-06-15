import json
import smtplib
from email.message import EmailMessage

import pandas as pd
import re


def extract_emails_from_excel(file_path):
    df = pd.read_excel(file_path)

    main_email = []
    for column in df.columns:
        for value in df[column]:
            if pd.notnull(value):
                email_matches = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', str(value))
                main_email.extend(email_matches)

    return main_email


def send_email(sender_email, receiver_email, email_subject, email_body, user_password):
    em = EmailMessage()
    em['From'] = sender_email
    em['To'] = receiver_email
    em['Subject'] = email_subject
    em.set_content(email_body)

    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(sender_email, user_password)
        smtp.send_message(em)


file_path = 'email.xlsx'
emails = extract_emails_from_excel(file_path)

file_path = 'body.txt'
with open(file_path, 'r') as file:
    file_contents = file.read()

file_path = 'subject.txt'
with open(file_path, 'r') as file:
    subject = file.read()

file_path = 'user.json'
with open(file_path, 'r') as file:
    main_details = json.load(file)

username = main_details['username']
password = main_details['password']

count = 0
for email in emails:
    body = 'Dear,\n' + file_contents
    send_email(username, email, subject, body, password)
    count += 1
    print(count)

