import email
import imaplib

import gspread
import win32com.client as win32
from oauth2client.service_account import ServiceAccountCredentials


def login_outlook():
    scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name("../json-secret-files/secret_key.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open('Email_List').sheet1

    # URL for Outlook connection
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 is the index of the Inbox folder

    # Define a filter to get emails with a specific subject and that are unread
    subject_filter = "[Subject]='New_Registration' AND [UnRead]=True"
    items = inbox.Items.Restrict(subject_filter)

    # Iterate through the messages and extract the subject and body
    for i, item in enumerate(items):
        sender = item.SenderEmailAddress
        subject = item.Subject
        body = item.Body
        datastr = body.split(",")
        Name = datastr[0]
        Address = datastr[1]
        NHSNO = datastr[2]
        PHNO = datastr[3]
        DOB = datastr[4]
        sex = datastr[5]
        status = "unsent"
        emailsend = datastr[6]
        print(Name, Address, NHSNO, PHNO, DOB, sex, status, emailsend)
        sheet.insert_row([Name, Address, NHSNO, PHNO, DOB, sex, status, emailsend], 2)

        # Mark the fetched emails as read
        item.UnRead = False


def login(username, password):

    scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
                 "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name("../json-secret-files/secret_key.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open('Email_List').sheet1

    # URL for IMAP connection
    imap_url = 'imap.gmail.com'

    # Connection with GMAIL using SSL
    my_mail = imaplib.IMAP4_SSL(imap_url)

    #  Log in using your credentials
    my_mail.login(username, password)

    # Select the Inbox to fetch unread messages
    my_mail.select('Inbox')

    # Define Key and Value for email search
    key = "SUBJECT"
    value = "New_Registration"
    _, data = my_mail.search(None, 'UNSEEN', key, value)  # Search for unread emails with specific key and value

    mail_id_list = data[0].split()  # IDs of all unread emails that we want to fetch

    msgs = []  # empty list to capture all messages
    # Iterate through messages and extract data into the msgs list
    for num in mail_id_list:
        typ, data = my_mail.fetch(num, '(RFC822)')  # RFC822 returns whole message (BODY fetches just body)
        msgs.append(data)

    # Iterate through the messages and extract the subject and body
    for i, msg in enumerate(msgs[::-1]):
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg = email.message_from_bytes((response_part[1]))
                From = my_msg['from']
                subject = my_msg['subject']
                body = ''
                for part in my_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload()
                datastr = body.split(",")
                Name = datastr[0]
                Address = datastr[1]
                NHSNO = datastr[2]
                PHNO = datastr[3]
                DOB = datastr[4]
                sex = datastr[5]
                status = "unsent"
                emailsend = datastr[6]
                print(Name, Address, NHSNO, PHNO, DOB, sex, status, emailsend)
                sheet.insert_row([Name, Address, NHSNO, PHNO, DOB, sex, status, emailsend], 2)

    for num in mail_id_list:
        my_mail.store(num, '+FLAGS', '\\Seen')
