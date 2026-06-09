import pandas as pd
import smtplib
import sys
sys.path.append(r"E:\python")
import DbConfig.imapconfig
from email.message import EmailMessage
import DbConfig.smtpconfig
from hrsheet import output_path
from email.utils import getaddresses
#from status import file_path
from excelsheetofmail import original_subject,original_message_id,original_from,original_cc
# Excel file path
#file_path = r"C:\Users\software.support\Desktop\HR_SHEET\output.xlsx"
file_path_mail=output_path
# Read Sheet3
df = (pd.read_excel(file_path_mail, sheet_name="Sheet3")).fillna('')

# Function to apply row colors
def highlight_status(row):

    if row["EMP_STATUS"] == 75:
        color = "background-color: #90EE90;"   # Light Green

    elif row["EMP_STATUS"] == 76:
        color = "background-color: #FFD580;"   # Light Orange

    else:
        color = ""

    # Return color only for visible columns
    return [color] * len(df_display.columns)
orange_count = (df['EMP_STATUS'] == 76).sum()
green_count = (df['EMP_STATUS'] == 75).sum()
#df_display = df.drop(columns=['EMP_STATUS'])
# Convert dataframe to styled HTML
# Remove EMP_STATUS column from display
df_display = df.drop(columns=['EMP_STATUS'])

# Function to apply row colors
def highlight_status(row):

    status = df.loc[row.name, "EMP_STATUS"]

    if status == 75:
        color = "background-color: #90EE90;"   # Green

    elif status == 76:
        color = "background-color: #FFD580;"   # Orange

    else:
        color = ""

    return [color] * len(row)

# Create HTML table
html_table = (
    df_display.style
    .apply(highlight_status, axis=1)
    .hide(axis="index")
    .set_table_styles([
        {
            'selector': 'table',
            'props': [
                ('border-collapse', 'collapse'),
                ('width', '100%')
            ]
        },
        {
            'selector': 'th',
            'props': [
                ('border', '1px solid black'),
                ('padding', '2px'),
                ('background-color', '#d9d9d9'),
                ('text-align', 'center'),
                ('font-size', '16px')
            ]
        },
        {
            'selector': 'td',
            'props': [
                ('border', '1px solid black'),
                ('padding', '2px'),
                ('text-align', 'center'),
                ('font-size', '16px')
            ]
        }
    ])
    .to_html()
)

# Create Email
# msg = EmailMessage()
# msg['Subject'] = "Exited Employee Details"
# msg['From'] = smtpconfig.SENDER_EMAIL

# # Multiple receivers support
# if isinstance(smtpconfig.RECEIVER_EMAILS, list):
#     msg['To'] = ", ".join(smtpconfig.RECEIVER_EMAILS)
# else:
#     msg['To'] = smtpconfig.RECEIVER_EMAILS
reply = EmailMessage()
reply['Subject'] = "Re: " + original_subject
reply['From'] = DbConfig.smtpconfig.SENDER_EMAIL
reply['To'] = original_from
#reply['To'] = smtpconfig.SENDER_EMAIL
cc_list = []

if original_cc:
    original_cc = original_cc.replace('\r', '').replace('\n', '').replace('\t', '')
    cc_list.append(original_cc)
cc_list.extend(DbConfig.smtpconfig.EXTRA_CC)

reply['Cc'] = ", ".join(cc_list)
reply['In-Reply-To'] = original_message_id
reply['References'] = original_message_id
# if original_cc:
#     parsed_cc = getaddresses([original_cc])

#     for name, addr in parsed_cc:
#         if addr:
#             cc_list.append(addr)

# cc_list.extend(smtpconfig.EXTRA_CC)

# reply['Cc'] = ", ".join(cc_list)
# Original CC from received mail
# if original_cc:
#     original_cc = original_cc.replace('\r', '').replace('\n', '').replace('\t', '')
#     cc_list.append(original_cc)

# # Add extra CC from config
# cc_list.extend(smtpconfig.EXTRA_CC)

# reply['Cc'] = ", ".join(cc_list)
# reply['In-Reply-To'] = original_message_id

# reply['References'] = original_message_id

# reply.add_alternative(f"""
# <html>
#     <body>

#         <p>Dear Team,</p>

#         <p>
#             {orange_count} employee’s LWD change
#             <span style="background-color:#FFD580;">
#                 (Orange marked)
#             </span>
#             and
#             {green_count} employee in-active
#             <span style="background-color:#90EE90;">
#                 (Green marked)
#             </span>,
#             PFB details :
#         </p>

#         <br>

#         {html_table}

#         <br>
#         <br>

#         <p>
#     Thanks & Regards,<br><br>
#     <b>Ravi Kumar Sharma</b> | Manager - Systems |<br>
#     Mobile : +91 9783699456<br>
#     Shalby Limited Corporate Office Ahmedabad<br>
#     <b><i>Go Green! Print this email only when necessary.</i></b>
# </p>

#     </body>
# </html>
# """, subtype='html')
# ================= EMAIL BODY =================
#<p>
#         <b>Color Legend:</b><br>
#       <span style="background-color:#90EE90;"> Green </span>= Employee Ststus Active<br>
#        <span style="background-color:#FFD580;"> Orange </span>= Employee Ststus In-Active
#         </p>

# Check if table has records
if df_display.empty:

    email_body = f"""
    <html>
    <body>

    <p>Dear Team,</p>

    <p>
    No discrepancy was found.
    The details in the sheet and SRIT remain the same. 
    </p>

    <br>

    <p>
    Thanks & Regards,<br><br>

    <b>Ravi Kumar Sharma</b> | Manager - Systems<br>
    Mobile: +91 9783699456<br>
    Shalby Limited Corporate Office Ahmedabad<br><br>

    <b><i>Go Green! Print this email only when necessary.</i></b>
    </p>

    </body>
    </html>
    """

else:

    email_body = f"""
    <html>
    <body>

    <p>Dear Team,</p>

    <p>
        {orange_count} employees LWD changed
        <span style="background-color:#FFD580;">
            (Orange Marked)
        </span>
        and
        {green_count} employees inactive
        <span style="background-color:#90EE90;">
            (Green Marked)
        </span>.
        Please find the details below:
    </p>

    <br>

    {html_table}

    <br>
    <p>
    <b>Color Legend:</b><br><br>
    <span style="background-color:#90EE90;"> Green </span>
    = Employee Status Active
    <br>
    <span style="background-color:#FFD580;"> Orange </span>
    = Employee Status Already Inactive
    </p>

    <br>

    <p>
    Thanks & Regards,<br><br>

    <b>Ravi Kumar Sharma</b> | Manager - Systems<br>
    Mobile: +91 9783699456<br>
    Shalby Limited Corporate Office Ahmedabad<br><br>

    <b><i>Go Green! Print this email only when necessary.</i></b>
    </p>

    </body>
    </html>
    """

reply.add_alternative(email_body, subtype='html')
# Send Email
try:
    with smtplib.SMTP(
        DbConfig.smtpconfig.SMTP_SERVER,
        DbConfig.smtpconfig.SMTP_PORT
    ) as server:

        server.starttls()

        server.login(
            DbConfig.smtpconfig.SENDER_EMAIL,
            DbConfig.smtpconfig.SENDER_PASSWORD
        )
        server.send_message(reply)

    DbConfig.imapconfig.imap.login(
    DbConfig.imapconfig.imap_EMAIL,
    DbConfig.imapconfig.imap_PASSWORD   
    )
    DbConfig.imapconfig.imap.append(
    "Sent",   # Folder name
    None,
    DbConfig.imapconfig.imaplib.Time2Internaldate(DbConfig.imapconfig.time.time()),
    reply.as_bytes()
    )
    DbConfig.imapconfig.imap.logout()
      
    #print("Email sent successfully!")

except Exception as e:
    print("Error while sending email:", e)