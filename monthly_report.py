import pymongo
import certifi
import pandas as pd
from datetime import datetime, timedelta
import smtplib
import os
import json
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()
MONGO_URI = os.getenv("MONGO_URI")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
ALL_MAIL = os.getenv("ALL_MAIL")
dept_mappings = json.loads(os.getenv("DEPT_MAPPINGS"))

client = pymongo.MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["studentDB"]
collection = db["latecomers"]

today = datetime.today()
first_day_last_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
last_day_last_month = (today.replace(day=1) - timedelta(days=1))

data = list(collection.find())
if not data:
    print("No data found. Skipping process.")
    exit()

df = pd.DataFrame(data)
df.drop(['_id', '__v'], axis=1, inplace=True, errors='ignore')
df['date'] = pd.to_datetime(df['date'])

df_month = df[(df['date'] >= first_day_last_month) & (df['date'] <= last_day_last_month)]
if df_month.empty:
    print("No records found for the last month. Skipping process.")
    exit()

df_month['date_str'] = df_month['date'].dt.strftime('%d/%m/%y')

student_counts = (
    df_month.groupby(['pin', 'name', 'department'])
    .agg(
        late_count=('date_str', lambda x: len(set(x))),
        repeated_dates=('date_str', lambda x: ', '.join(sorted(set(x))))
    )
    .reset_index()
)

df_filtered = student_counts[student_counts['late_count'] >= 5]
if df_filtered.empty:
    print("No students were late on 5 or more unique days. Skipping process.")
    exit()

month_year = first_day_last_month.strftime('%B-%Y')
consolidated_filename = f"Monthly_Latecomers_{month_year}.xlsx"
saved_files = {}

with pd.ExcelWriter(consolidated_filename, engine="xlsxwriter") as writer:
    for dept, email in dept_mappings.items():
        df_dept = df_filtered[df_filtered['department'] == dept]
        if not df_dept.empty:
            dept_filename = f"{dept}_monthly_latecomers_{month_year}.xlsx"
            df_dept.to_excel(writer, sheet_name=dept, index=False)
            df_dept.to_excel(dept_filename, index=False)
            saved_files[dept] = dept_filename
            print(f"Saved: {dept_filename}")

    print(f"Consolidated report saved: {consolidated_filename}")

def send_email(receiver_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as file:
        msg.add_attachment(file.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=os.path.basename(attachment_path))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        print(f"Email sent to {receiver_email} with {attachment_path}")

    os.remove(attachment_path)
    print(f"Deleted: {attachment_path}")

for dept, email in dept_mappings.items():
    if email and dept in saved_files:
        send_email(email, f"Monthly Latecomers Report - {dept} ({month_year})",
                   "Attached is the list of students who were late on 5 or more unique days last month.", saved_files[dept])

send_email(ALL_MAIL, f"Monthly Latecomers Consolidated Report ({month_year})",
           "Attached is the consolidated latecomers' report for all departments last month.", consolidated_filename)
