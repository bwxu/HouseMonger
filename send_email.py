import imaplib
import time
import email.message
import xlrd

# call before running anything else
# opens the gmail draft mailbox and returns a conn which can be used to create new messages 
def initialize():
    conn = imaplib.IMAP4_SSL('imap.gmail.com')
    conn.login('plphousemonger@gmail.com', 'pilambdaphi')
    conn.select('[Gmail]/Drafts')
    return conn


# msg is a valid email message
# creates a draft of the email in the gmail inbox
def create_email(msg, conn):
    now = imaplib.Time2Internaldate(time.time())
    conn.append('[Gmail]/Drafts', '', now, str(msg))


# date is a string in YYYY-MM-DD format
# recipient is a string that is a valid email address
# conn is the object obtained from intialize()
# sends trash duty reminder the day before the scheduled date
def trash_reminder(date, recipient, floor, conn):
    msg = email.message.Message()
    msg['Subject'] = '[House-Monger] Floor ' + str(floor) + ' Trash Duty Reminder @' + date + '(-1)'
    msg['From'] = 'plphousemonger@gmail.com'
    msg['To'] = recipient
    msg.set_payload('This a reminder for your upcoming trash duty on ' + date + '. ' + 
                    'Please complete this task by the end of the day tomorrow. ' + 
                    'Don\'t forget to send a monger a photo after you\'re done. ')
    create_email(msg, conn)


# date is a string in YYYY-MM-DD format
# recipients is a list of strings that are valid email addresses
# conn is the object obtained from initialize()
# sends dinner duty reminder the day before the scheduled date
def dinner_duty_reminder(date, recipients, conn):
    msg = email.message.Message()
    msg['Subject'] = '[House-Monger] Dinner Duty Reminder @' + date + '(-1)'
    msg['From'] = 'plphousemonger@gmail.com'
    msg['To'] = '; '.join(filter(None, recipients))
    msg.set_payload('This a reminder for your upcoming dinner duty tomorrow on ' + date + '. ' + 
                    'Don\'t forget to send a monger a photo after you\'re done. ')
    create_email(msg, conn)
    
# date is a string in YYYY-MM-DD format representing the shazam start date
# recipients is a list of strings that are valid email addresses
# conn is the object obtained from initialize()
# shazam is the string name of the shazam
# sends recurring shazam reminder the day before the scheduled date
def shazam_reminder(date, recipients, conn, shazam):
    msg = email.message.Message()
    msg['Subject'] = '[House-Monger] Shazam Reminder: ' + shazam + ' @' + date + ' repeat every week'
    msg['From'] = 'plphousemonger@gmail.com'
    msg['To'] = '; '.join(filter(None, recipients))
    msg.set_payload('This a reminder for your upcoming shazam, ' + shazam + 
                    ', which is due by the end of Sunday this weekend. ' +
                    ' Please take care to complete everything in the shazam description. ' +
                    'Don\'t forget to send a monger a photo after you\'re done. ')
    create_email(msg, conn)

def create_dinner_duty_emails(sheet, conn):
    dinner = xlrd.open_workbook(sheet).sheet_by_index(0)
    for row in range(1, dinner.nrows):
        if dinner.cell(row, 0).value != '':
            dinner_duty_reminder(dinner.cell(row, 0).value, [dinner.cell(row, i).value for i in range(4, 7)], conn)


def create_shazam_emails(sheet, conn):
    shazam = xlrd.open_workbook(sheet).sheet_by_index(0)
    start_date = '2016-10-14'
    for row in range(1, shazam.nrows):
        shazam_reminder(start_date, [shazam.cell(row, i).value for i in range(5, 8)], conn,
                        shazam.cell(row, 0).value + ' - ' + shazam.cell(row, 1).value)

    
def create_trash_duty_emails(sheet, floor, conn):
    trash = xlrd.open_workbook(sheet).sheet_by_index(0)
    for row in range(1, trash.nrows):
        trash_reminder(trash.cell(row, 0).value, trash.cell(row, 1).value, floor, conn)

        
conn = initialize()
dinner = xlrd.open_workbook('dinner_duty.xlsx').sheet_by_index(0)
shazam = xlrd.open_workbook('shazams.xlsx').sheet_by_index(0)
trash2 = xlrd.open_workbook('trash2.xlsx').sheet_by_index(0)
trash3 = xlrd.open_workbook('trash3.xlsx').sheet_by_index(0)
trash4 = xlrd.open_workbook('trash4.xlsx').sheet_by_index(0)

create_trash_duty_emails('trash2.xlsx', 2, conn)
create_trash_duty_emails('trash3.xlsx', 3, conn)
create_trash_duty_emails('trash4.xlsx', 4, conn)

#create_dinner_duty_emails('dinner_duty.xlsx', conn)
#create_shazam_emails('shazams.xlsx', conn)