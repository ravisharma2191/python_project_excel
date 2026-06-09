import imaplib
import email
import sys
sys.path.append(r"E:\python")
import os
from datetime import datetime
from email.utils import parsedate_to_datetime, parseaddr
from DbConfig.smtpconfig import SENDER_EMAIL_2, SENDER_PASSWORD

# ---------------- CONFIG ----------------
IMAP_SERVER = "shalbyhospitals.icewarpcloud.in"
SENDER_FILTER = "exe1.hr@shalby.in"

SAVE_FOLDER = r"E:\python\Project_1\HR_SHEET"
os.makedirs(SAVE_FOLDER, exist_ok=True)

today = datetime.now().date()
today_str = datetime.now().strftime("%d-%b-%Y")

# ---------------- CONNECT ----------------
mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
mail.login(SENDER_EMAIL_2, SENDER_PASSWORD)
mail.select("INBOX")

# ✅ Filter from IMAP itself
status, messages = mail.search(None, f'(FROM "{SENDER_FILTER}" SINCE "{today_str}")')
email_ids = messages[0].split()

for e_id in email_ids:
    status, msg_data = mail.fetch(e_id, "(RFC822)")

    for response_part in msg_data:
        if isinstance(response_part, tuple):

            msg = email.message_from_bytes(response_part[1])
            original_message_id = msg.get("Message-ID")
            original_subject = msg.get("Subject")
            original_from = msg.get("From")
            original_cc = msg.get("Cc")

            #print("Message-ID:", original_message_id)

            sender = msg.get("From")
            date_header = msg.get("Date")

            if not date_header or not sender:
                continue

            try:
                msg_date = parsedate_to_datetime(date_header).date()
            except Exception:
                continue

            # Extract clean email
            email_address = parseaddr(sender)[1]

            if msg_date == today and email_address.lower() == SENDER_FILTER.lower():

                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":

                        filename = part.get_filename()

                        if filename and filename.lower().endswith(".xlsx"):

                            # ✅ Add timestamp
                            #timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"{filename}"

                            file_path = os.path.join(SAVE_FOLDER, filename)

                            with open(file_path, "wb") as f:
                                f.write(part.get_payload(decode=True))

                                # cc_list = []

                                # if original_cc:
                                #     original_cc = original_cc.replace('\r', '').replace('\n', '').replace('\t', '')
                                #     cc_list.append(original_cc)

                            #print("original_cc:", original_message_id)

mail.logout()