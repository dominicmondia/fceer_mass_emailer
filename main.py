import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2 import service_account
import json
import smtplib


load_dotenv()
GOOGLE_CREDS = os.getenv("GOOGLE_CREDS")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SHEET_URL = os.getenv("SHEET_URL")


def load_recipients(creds: str, sheet_url: str) -> pd.DataFrame:
    service_account_info = json.loads(creds)
    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = credentials.with_scopes(scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url)

    data = sheet.get_worksheet(0).get_all_records()
    return pd.DataFrame.from_dict(data)


recipient_data = load_recipients(GOOGLE_CREDS, SHEET_URL)

for recipient in recipient_data.itertuples():
    first_name = recipient[1].strip()
    email_address = recipient[3]
    with smtplib.SMTP("smtp.gmail.com", port=587) as connection:
        connection.starttls()
        connection.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        connection.sendmail(
            from_addr=EMAIL_ADDRESS,
            to_addrs=email_address,
            msg=f""
                f"Subject:FCEER 2024\n\n"
                f"Hi {first_name},\n\n"
                f"We are pleased to inform you that you are one of the selected few granted the opportunity to reserve "
                f"a slot for this year's Free College Entrance Exam Review.\n"
                f"Kindly respond to this message to confirm your intention to proceed "
                f"or opt out of the program BEFORE 11:59 PM\n\n"
                f"Best regards,\n"
                f"FCEER Guild"
        )

