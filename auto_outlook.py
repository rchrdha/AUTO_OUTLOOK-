import smtplib, os
import pandas as pd
from email.message import EmailMessage 

# get user login data for Outlook 
from_address = input('Enter Email: ')
password = input('Enter Password: ')

# get data from xlsx, checks for inconsistencies
try:
    mailing_list = pd.read_excel(os.path.expanduser("~/Desktop/FILENAME.xlsx"), sheet_name = "Sheet1")
    all_clubnames = mailing_list["Name"]
    all_clubemails = mailing_list["Email"]
except KeyError:
    print("Something went wrong... Check your xlsx column titles for extra whitespaces.")
else: 
    mailing_list = pd.read_excel(os.path.expanduser("~/Desktop/FILENAME.xlsx"), sheet_name = "Sheet1")
    all_clubnames = mailing_list["Name"]
    all_clubemails = mailing_list["Email"]


# connect to server and login 
with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
    server.connect("smtp-mail.outlook.com", 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(from_address, password)

    # email customisation and sending
    for club, club_email in zip(all_clubnames, all_clubemails):
        msg = EmailMessage() 
        msg['Subject'] = 'totally real email subject' 
        msg['From'] = from_address
        msg['To'] = {club_email}
        msg.set_content(f'This is a test message body for {club}')
        server.send_message(msg)
        print(f'Email successfully sent for {club} at {club_email}')

    server.quit()
