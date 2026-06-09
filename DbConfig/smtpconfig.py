import smtplib
from email.mime.text import MIMEText

SMTP_SERVER = "shalbyhospitals.icewarpcloud.in"   # change if internal SMTP
SMTP_PORT = 587

SENDER_EMAIL = "Ravi Sharma <srmgr.systems@shalby.org>"
SENDER_EMAIL_2 = "srmgr.systems@shalby.org"
SENDER_PASSWORD = "Shalby@2191"   # or app password

RECEIVER_EMAILS = [
    "Ravi Sharma <srmgr.systems@shalby.org>"
]
EXTRA_CC= ["Kaptan Kishore <cio@shalby.org>","Jigar Shah <headitsoftware@shalby.org>"]