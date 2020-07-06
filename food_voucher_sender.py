from email.mime.text import MIMEText
from datetime import timedelta
import smtplib
import time
import xlrd

# Office 365 imposes a limit of
# 30 messages sent per minute, and a limit of 10,000 recipients per day.

def openData():
    # Open Excel document with urls
    f = ('codes.xlsx')
    wb = xlrd.open_workbook(f)
    dataSheet = wb.sheet_by_index(0)
    return(getValues(dataSheet))

def openContacts():
    # Open Excel document with contacts
    f = ('contacts.xlsx')
    wb = xlrd.open_workbook(f)
    emailSheet = wb.sheet_by_index(0)
    return(emailSheet)

def getValues(dataSheet):
    # Extract family ID and URLS, add to dictionary of lists. 
    # e.g. {1:['URL'], 2:['URL','URL']}
    amt_urls = dataSheet.nrows-1
    count = 1
    stored = ''
    id_urls = {}
    for i in range(amt_urls):
        url = dataSheet.cell_value(count,6)
        fID = dataSheet.cell_value(count,0)
        if fID == '': # If blank, use previous ID
            fID = stored
        fID = int(fID) # Convert to int for niceness
        if fID in id_urls: 
            id_urls[fID].append(url) # If ID already exists, append new url so one email for multiple links
        else:
            id_urls[fID] = [url]          
        stored = fID # Stored ID incase next blank therefore same family
        count+=1
    return(sendEmail(id_urls))

def sendEmail(id_urls):
    # Create email and send to respective family.
    # Connect constants
    svr = 'smtp.office365.com'
    port = '587'
    sender = '' # Email address e.g. test@outlook.com
    pwd = ''    # Password for above email
    # Email constant
    subject = 'Voucher Codes' 
    
    emailSheet = openContacts()
    count = 1
    vcTot = 0
    start = time.time()
    for i in id_urls:
        vc = ''
        email = emailSheet.cell_value(count,1)
        code = id_urls.get(i, '')
        for i in code:
            vcTot += 1
            vc = str(vc)+str(i)+'\n\n'
        # Format email
        body = f'EMAIL BODY HERE \n\n Code: {vc}'

        msg = MIMEText(body)
        msg['To'] = email
        msg['From'] = sender
        msg['Subject'] = subject
        # print(msg) # DEBUG
        # print("*"*20) # DEBUG 
        # Send it
        server = smtplib.SMTP(svr, port)
        server.ehlo()
        server.starttls()
        server.login(sender, pwd)
        server.send_message(msg)
        print(f"Email sent to: {email} with {len(code)} code(s)")
        print("*"*20)
        server.quit()
        count += 1
    end = str(timedelta(seconds=round(time.time() - start, 2)))[2:10]
    print(f'{vcTot} codes sent to {count-1} emails in {end}')
    return()
        

def main():
    return(openData())

if __name__ == "__main__":
    main()