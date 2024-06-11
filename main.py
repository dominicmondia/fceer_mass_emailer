import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2 import service_account
import json

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application

load_dotenv()
GOOGLE_CREDS = os.getenv("GOOGLE_CREDS")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SHEET_URL = os.getenv("SHEET_URL")


def load_recipients(creds: str, sheet_url: str) -> pd.DataFrame:
    service_account_info = json.loads(creds)
    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = credentials.with_scopes(scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url)

    data = sheet.get_worksheet(0).get_all_records()
    return pd.DataFrame.from_dict(data)


def update_email_status(client, sheet_url, row, status):
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    worksheet.update_cell(row + 2, worksheet.find("Email Status").col, status)  # row + 2 to account for header row


def send_emails(recipient_data: pd.DataFrame):
    recipient_data = load_recipients(GOOGLE_CREDS, SHEET_URL)
    fb_group_link = "https://facebook.com/groups/1137865990833763/"
    body_template = """
    <html>
      <body>
        <p>Hi {first_name},</p>
        <p>We are pleased to inform you that you are one of the selected few granted the opportunity to reserve a slot for this year's Free College Entrance Exam Review!</p>
        <p>If you wish to decline this reservation, you may respond to this email stating your decision to opt out.</p>
        <p>However, if you wish to proceed with the program, kindly perform the following steps:</p>
        <ol>
          <li>Respond to this message confirming your interest to commit with the review before June 13 11:59 PM</li>
          <li>Download the waiver attached to this email and have it signed</li>
          <li>Join the private FB group linked <a href="{fb_group_link}">here</a> for FCEER 19 for further updates and instructions</li>
        </ol>
        <p>We hope to welcome you to a productive summer.</p>
        <p>Sincerely,<br>
        SJDM Free College Entrance Examination Guild</p>
      </body>
    </html>
    """

    filename = 'waiver.pdf'
    with open(filename, 'rb') as fp:
        att = email.mime.application.MIMEApplication(fp.read(), _subtype="pdf")
    att.add_header('Content-Disposition', 'attachment', filename=filename)

    service_account_info = json.loads(GOOGLE_CREDS)
    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = credentials.with_scopes(scope)
    client = gspread.authorize(creds)

    for index, recipient in recipient_data.iterrows():
        first_name = recipient['First Name'].title().strip()
        email_address = recipient['Email Address']

        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = email_address
        msg["Subject"] = "FCEER 2024 Slot Reservation"

        body = body_template.format(first_name=first_name, fb_group_link=fb_group_link)
        msg.attach(MIMEText(body, 'html'))
        msg.attach(att)

        try:
            with smtplib.SMTP("smtp.gmail.com", port=587) as connection:
                connection.starttls()
                connection.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                connection.sendmail(
                    from_addr=EMAIL_ADDRESS,
                    to_addrs=email_address,
                    msg=msg.as_string()
                )
            update_email_status(client, SHEET_URL, index, "Sent")
        except Exception as e:
            print(f"Error sending email to {email_address}: {e}")
            update_email_status(client, SHEET_URL, index, "Failed")


if __name__ == "__main__":
    recipients = load_recipients(GOOGLE_CREDS, SHEET_URL)
    send_emails(recipients)
