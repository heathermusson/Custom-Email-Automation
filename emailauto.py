#Created by Heather Musson on May 14, 2018 (Last updated May 30, 2018)

import smtplib
import xlrd #package to read excel file  - xlwt to write files
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

#get contacts
def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding = 'utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails

#get contacts from an excel file
def get_contacts_excel(filename, sheetname):
    names = []
    emails = []
    #Opens excel file
    book = xlrd.open_workbook(filename)

    #Gets sheet
    sheet = book.sheet_by_name(sheetname)
    print (sheet.nrows)

    for i in range(sheet.nrows):
        row = sheet.row_values(i)
        names.append(row[0])
        emails.append(row[1])

    return names, emails

#import message
def read_template(filename):
    with open(filename, 'r', encoding = 'utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():

    #set up smtp
    try:
        s = smtplib.SMTP_SSL(host='smtp.gmail.com', port = '465')
        s.ehlo()
        s.login('email', 'password')
    except:
        print ("Error in setting up email connection")

    #import contacts
    #names, emails = get_contacts('mycontacts.txt')
    names, emails = get_contacts_excel('contacts.xlsx', 'Sheet1')

    #import message template
    message_template_html = read_template('messagehtml.txt')
    message_template_plain = read_template('messageplain.txt')

    #send email for each contact
    for name, email in zip(names, emails):
        msg = MIMEMultipart() #creates the message

        #put customized name in message
        messagehtml = message_template_html.substitute(PERSON_NAME = name.title())
        messageplain = message_template_plain.substitute(PERSON_NAME = name.title())

        #prints message body
        print(messageplain)

        msg['From'] = 'emails'
        msg['To'] = email
        msg['Subject'] = "This is a test"

        #Encapsulate the plain and HTML versions of the message body in an 'alternative' part
        #Message agents can decide which they want to display

        #Plain-text option
        msgAlternative = MIMEMultipart('alternative')
        msg.attach(msgAlternative)
        msgText = MIMEText(messageplain, 'plain')
        msgAlternative.attach(msgText)

        #HTML option
        msgText = MIMEText(messagehtml, 'html')
        msgAlternative.attach(msgText)

        #Assumes the image is in the current directory
        fp = open('image.jpg', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        #Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<image1>')
        msg.attach(msgImage)

        s.send_message(msg)
        del msg
        print("Message Sent to " + email)

    s.close()

if __name__ == '__main__':
    main()
