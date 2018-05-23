#Created by Heather Musson on May 14, 2018 (Last updated May 23, 2018)

import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_EMAIL = 'email'
MY_PASSWORD = 'password'

#get contacts
def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding = 'utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
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
        s.helo()
        s.login(MY_EMAIL, MY_PASSWORD)
    except:
        print ("Error in setting up email connection")

    #import contacts
    names, emails = get_contacts('mycontacts.txt')

    #import message template
    message_template = read_template('message.txt')

    #send email for each contact
    for name, email in zip(names, emails):
        msg = MIMEMultipart() #creates the message

        #put customized name in message
        message = message_template.substitute(PERSON_NAME = name.title())

        #prints message body
        print(message)

        msg['From'] = MY_EMAIL
        msg['To'] = email
        msg['Subject'] = "This is a test"
        msg.attach(MIMEText(message, 'plain')) #adds message body
        s.sendmail(MY_EMAIL, email, msg)
        del msg
        print("Message Sent")

    s.close()

if __name__ == '__main__':
    main()
