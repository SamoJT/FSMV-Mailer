from email.mime.text import MIMEText
from datetime import timedelta
import smtplib
import time
import xlrd
###################################################
# TO DO:
# Include rate limit check / catch to prevent code exiting due to hitting limit
###################################################

# Office 365 imposes a limit of
# 30 messages sent per minute, and a limit of 10,000 recipients per day.

def openData():
    # Open Excel document with urls
    f = ('vouchers.xlsx')
    wb = xlrd.open_workbook(f)
    dataSheet = wb.sheet_by_index(0)
    return(getValues(dataSheet))

def getValues(dataSheet):
    # Extract email and codes, add to dictionary of email:[code] pairs. 
    # e.g. {a@gmail.com:['CODE'], b@yahoo.co.uk:['CODE','CODE']}
    amt_urls = dataSheet.nrows-1
    count = 1
    stored = ''
    email_urls = {}
    emailColumn = 3
    codeColumn = 7
    if dataSheet.cell_value(0,codeColumn) != 'Code' or dataSheet.cell_value(0,emailColumn) != 'Email':
        return("ERROR - Either Email or Code column mismatched.")
    for i in range(amt_urls):
        code = dataSheet.cell_value(count,codeColumn) # cell_value(ROW,COL)
        eAddr = dataSheet.cell_value(count,emailColumn)
        if eAddr == '': # If blank, use previous Email
            eAddr = stored
        if eAddr in email_urls: 
            email_urls[eAddr].append(code) # If email already exists, append new code so one email for multiple codes
        else:
            email_urls[eAddr] = [code]          
        stored = eAddr # Stored email incase next blank therefore same family
        count+=1
    return(sendEmail(email_urls))

def sendEmail(email_urls):
    # Create email and send to respective family.
    # Connect constants
    svr = 'smtp.office365.com'
    port = '587'
    sender = '' # Outlook Email address here. e.g. test@outlook.com
    pwd = ''    # Plaintext password here
    # Email constant
    subject = 'Voucher Codes' 
    
    count = 1
    vcTot = 0
    start = time.time()
    for i in email_urls:
        vc = ''
        email = i
        code = email_urls.get(i, '')
        for i in code:
            vcTot += 1
            vc = str(vc)+str(i)+'\n\n'
        # Format email
        body = f'EMAIL BODY HERE\n\nCodes: {vc}'

        msg = MIMEText(body)
        msg['To'] = email
        msg['From'] = sender
        msg['Subject'] = subject
        # print(msg) # DEBUG
        # print("*"*20) # DEBUG 
        # Send it #
        server = smtplib.SMTP(svr, port) #
        server.ehlo() #
        server.starttls() #
        server.login(sender, pwd) #
        server.send_message(msg) #
        print(f"Email sent to: {email} with {len(code)} code(s)")
        print("*"*20)
        server.quit() #
        count += 1
    end = str(timedelta(seconds=round(time.time() - start, 2)))[2:10]
    print(f'{vcTot} codes sent to {count-1} emails in {end}')
    return()
        

def main():
    return(openData())

if __name__ == "__main__":
    main()