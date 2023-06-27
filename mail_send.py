import smtplib

import gspread
import win32com.client as win32
from oauth2client.service_account import ServiceAccountCredentials


def outlook_mail_send(self):
    scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name("../json-secret-files/secret_key.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open('Email_List').sheet1

    # Get all the values in the first column
    column_values = sheet.col_values(1)

    # Count the number of non-empty rows
    num_rows = len([value for value in column_values if value])

    print(num_rows)

    for i in range(2, num_rows + 1):
        if sheet.cell(i, 7).value == 'unsent':
            olApp = win32.Dispatch('Outlook.Application')
            OINS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)

            mailItem.Subject = 'Registration Successfully'
            mailItem.BodyFormat = 1
            name = sheet.cell(i, 1).value
            mailItem.Body = "Dear " + name + ",\n\n Congratulations!! \n Thank you so much for registering our service. \n\n\n Thanks and Regards, \n NHS Team"
            mailItem.Sender = 'udaysankar.mukherjee2021@iem.edu.in'
            mailItem.To = sheet.cell(i, 8).value

            mailItem.Display()
            mailItem.Save()
            mailItem.Send()
            sheet.update_cell(i, 7, "sent")


def google_mail_send(self):
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
             'https://www.googleapis.com/auth/drive']

    creds = ServiceAccountCredentials.from_json_keyfile_name('../json-secret-files/secret_key.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('Email_List').sheet1

    # Get all the values in the first column
    column_values = sheet.col_values(1)

    # Count the number of non-empty rows
    num_rows = len([value for value in column_values if value])

    for i in range(2, num_rows + 1):
        if sheet.cell(i, 7).value == 'unsent':
            # Connect to Gmail SMTP server
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.ehlo()
            server.starttls()

            # Login to Gmail account
            gmail_user = 'udaysankar2003@gmail.com'
            gmail_password = 'yxoqclmsdmqsjxeg'
            server.login(gmail_user, gmail_password)

            # Compose the message
            name = sheet.cell(i, 1).value
            subject = 'Registration Successfully'
            body = f"Dear {name},\n\nCongratulations!!\nThank you so much for registering our service.\n\n\nThanks and Regards,\nNHS Team"
            message = f'Subject: {subject}\n\n{body}'

            # Send the email
            to_address = sheet.cell(i, 8).value
            server.sendmail(gmail_user, to_address, message)

            # Mark the email as sent in the Google Sheets file
            sheet.update_cell(i, 7, 'sent')

            # Close the connection to the SMTP server
            server.quit()