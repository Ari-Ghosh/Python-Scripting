import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pywinauto.application import Application

import win32com.client as win32

scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("../json-secret-files/secret_key.json", scope)
client = gspread.authorize(creds)
sheet = client.open('Email_List').sheet1

# Get all the values in the first column
column_values = sheet.col_values(1)

# Count the number of non-empty rows
num_rows = len([value for value in column_values if value])

for i in range(2, num_rows + 1):
    if sheet.cell(i, 7).value == 'unsent':
        name = sheet.cell(i, 1).value
        address = sheet.cell(i, 2).value
        nhs_no = sheet.cell(i, 3).value
        ph_no = sheet.cell(i, 4).value
        dob = sheet.cell(i, 5).value
        sex = sheet.cell(i, 6).value
        email = sheet.cell(i, 8).value
        poaddress = sheet.cell(i, 9).value
        idproof = sheet.cell(i, 10).value
        # time.sleep(5)
        # app.MultiStepForm.print_control_identifiers()

        app = Application(backend='uia').start("../Form/Multi_Step_Form Setup 1.0.0.exe")
        app = Application(backend='uia').connect(title='Multi_Step_Form', timeout=10)

        nameEditor = app.MultiStepForm.child_window(title="Name", control_type="Edit").wrapper_object()
        nameEditor.type_keys(name, with_spaces=True)

        emailEditor = app.MultiStepForm.child_window(title="Email Address", control_type="Edit").wrapper_object()
        emailEditor.type_keys(email, with_spaces=True)

        phoneEditor = app.MultiStepForm.child_window(title="Phone Number", control_type="Edit").wrapper_object()
        phoneEditor.type_keys(ph_no)

        nextbutton = app.MultiStepForm.child_window(title="Next", control_type="Button").wrapper_object()
        nextbutton.click_input()

        nhsEditor = app.MultiStepForm.child_window(title="Enter 10 Digit NHS No. (Optional)",
                                                   control_type="Edit").wrapper_object()
        nhsEditor.type_keys(nhs_no, with_spaces=True)

        nextbutton = app.MultiStepForm.child_window(title="Next", control_type="Button").wrapper_object()
        nextbutton.click_input()
        app.MultiStepForm.print_control_identifiers()
        pronoun = ''
        if sex == 'Male':
            pronoun = 'he/him'
            female_checkbox = app.MultiStepForm.child_window(auto_id="male", control_type="CheckBox").wrapper_object()
            female_checkbox.click_input()

        else:
            pronoun = 'she/her'
            female_checkbox = app.MultiStepForm.child_window(auto_id="female", control_type="CheckBox").wrapper_object()
            female_checkbox.click_input()

        pronounEditor = app.MultiStepForm.child_window(title="Pronoun", control_type="Edit").wrapper_object()
        pronounEditor.type_keys(pronoun, with_spaces=True)
        dobEditor = app.MultiStepForm.child_window(title="Date of Birth", control_type="Edit").wrapper_object()
        dobEditor.type_keys(dob)
        addressEditor = app.MultiStepForm.child_window(title="Permanent Address", control_type="Edit").wrapper_object()
        addressEditor.type_keys("newtown , kolkata", with_spaces=True)
        nextbutton = app.MultiStepForm.child_window(title="Next", control_type="Button").wrapper_object()
        nextbutton.click_input()

        poaddressEditor = app.MultiStepForm.child_window(title="Proof of Address", control_type="Edit").wrapper_object()
        poaddressEditor.type_keys(poaddress, with_spaces=True)

        idproof = app.MultiStepForm.child_window(title="Identity Proof", control_type="Edit").wrapper_object()
        idproof.type_keys(idproof, with_spaces=True)

        submitbutton = app.MultiStepForm.child_window(title="Submit", control_type="Button").wrapper_object()
        submitbutton.click_input()

        closebutton = app.MultiStepForm.child_window(title="Close", control_type="Button").wrapper_object()
        closebutton.click_input()

        olApp = win32.Dispatch('Outlook.Application')
        OINS = olApp.GetNameSpace('MAPI')

        mailItem = olApp.CreateItem(0)

        mailItem.Subject = 'Registration Successfully'
        mailItem.BodyFormat = 1
        name = sheet.cell(i, 1).value
        mailItem.Body = "Dear " + name + ",\n\n Congratulations!! \n Thank you so much for registering our service. \n\n\n Thanks and Regards, \n NHS Team"
        mailItem.Sender = 'Arijit.Ghosh2021@iem.edu.in'
        mailItem.To = sheet.cell(i, 8).value

        mailItem.Display()
        mailItem.Save()
        mailItem.Send()
        sheet.update_cell(i, 7, "sent")
