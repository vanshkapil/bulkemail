import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
import time

def mail(email,body, subject):
    server = smtplib.SMTP('smtp-mail.outlook.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    Subject = subject
    user = USERNAME
    password = PASSWORD
    to_addr = email

    Body = body
    server.login(user,password)

    msg = f"Subject: {Subject}\n\n {Body}"

    # msg = MIMEMultipart('mixed')
    # msg['Subject'] = subject
    # msg['From'] = user
    # msg['To'] = email
    # msg['Disposition-Notification-To'] = user



    server.sendmail(
        user,
        to_addr,
        msg,
        rcpt_options=['NOTIFY=SUCCESS,DELAY,FAILURE']
    )

    print('Mail sent', msg)
    server.quit()

def gsheetAuth():
    """
    Authenticates google sheets.
    :return: sheet obj
    """
    # required by google sheets. service_account.json should in in same directory as this file
    gc = gspread.service_account(filename=GOOGLEAPISERVICE_ACCOUNT)
    # url is the google sheet url
    sht2 = gc.open_by_url(<GOOGLE_SPREADSHEET_URL>)
    return sht2


sheet= "email_list"
sht2= gsheetAuth()
companySheet = sht2.worksheet(sheet)
rec = companySheet.get_all_records(head=1)

for record in range(len(rec)):
    print('record is ',record)
    if rec[record]['sent_status']=='':
        email = str(rec[record]['Email_id'])
        print(type(email))
        subject = "This is a test email"
        body = "Test email body"
        mail(email,body,subject)
        print('Mail sent to ',email,' now sleeping...')
        companySheet.update_cell(record+2,2,'sent')
        time.sleep(60)



