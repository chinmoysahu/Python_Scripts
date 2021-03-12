import smtplib

mailserver = smtplib.SMTP('email-smtp.us-east-1.amazonaws.com',465)
mailserver.ehlo()
mailserver.starttls()
mailserver.login('U', 'P')

toaddr = 'c@gmail.com'
cc = ['c@gmail.com','c@gmail.com']
bcc = ['c@gmail.com']

#toaddrs = [toaddr] + cc + bcc

mailserver.sendmail('c@gmail.com',[toaddr,cc],'Subject: Test\n\nDear Chinmoy, \nReminder to submit Smart track for Weekending xxxx/xx/xx.\nThis is Timesheet Default Reminder#2 in this month and Reminder #4 YTD. ')
mailserver.quit()
