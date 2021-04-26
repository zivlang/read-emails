import email
import re
import imaplib
from datetime import datetime, date
import smtplib

email_user = 'my@address.com'
email_pass = 'mypassword'
receivers = ['receiving@address.com']
#  smtp
smtpHost = 'smtp.office365.com'
smtpPort = 587

M = imaplib.IMAP4_SSL('imap-mail.outlook.com', 993)
M.login(email_user, email_pass)
M.select('inbox')
typ, message_numbers = M.search(None, 'OR OR OR OR OR SUBJECT bioinformatic SUBJECT bioinformatics '
                                      'SUBJECT bioinformatician SUBJECT python SUBJECT "big data" '
                                      'SUBJECT "data science"')


def send_email(content):
    # Add the From: and To: headers at the start!
    message = f"From: {email_user}\r\nTo: {','.join(receivers)}\r\nSubject: from {content} \r\nDate: \r\n"
    try:
        smtpObj = smtplib.SMTP(smtpHost, smtpPort)
        # smtpObj.set_debuglevel(1)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.ehlo()
        smtpObj.login(email_user, email_pass)
        smtpObj.sendmail(email_user, receivers, message)
        smtpObj.quit()
        print("Successfully sent email")
    except smtplib.SMTPException as e:
        print(e)
        print("Error: unable to send email")


for num in message_numbers[0].split()[-5:]:
    typ, data = M.fetch(num, '(RFC822)')
    d = M.fetch(num, '(BODY[HEADER.FIELDS (SUBJECT)])')
    # num1 = base64.b64decode(num)          # unnecessary, I think
    email_content = data[0][1].decode('utf-8')
    if 'seminar' in email_content:
        continue
    email_date = re.search('Date: .{5}(.*?).{6}\r\n', email_content).group(1)
    email_content = re.sub('Subject: ', f'Subject: (Sent on {email_date}) ', email_content)
    send_email(email_content)

M.close()
M.logout()
