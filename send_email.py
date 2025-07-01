import os
import smtplib
from email.message import EmailMessage

SMTP_USER = os.environ["SENDER_EMAIL"]
SMTP_PASS = os.environ["GMAIL_APP_PASSWORD"]
TO_EMAIL = os.environ["RECIPIENT_EMAIL"]

# Email content
msg = EmailMessage()
msg["Subject"] = "Daily API Report"
msg["From"] = SMTP_USER
msg["To"] = TO_EMAIL
msg.set_content("Here is your daily report.")

# Optional: attach a file if exists
filename = "report.csv"
if os.path.exists(filename):
    with open(filename, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=filename)

# Send email
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(SMTP_USER, SMTP_PASS)
    smtp.send_message(msg)

print("✅ Email sent.")
