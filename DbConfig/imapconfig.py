import imaplib
import time

IMAP_SERVER = "shalbyhospitals.icewarpcloud.in"
IMAP_PORT = 993

imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)

imap_EMAIL = "srmgr.systems@shalby.org"
imap_PASSWORD = "Shalby@2191" 

