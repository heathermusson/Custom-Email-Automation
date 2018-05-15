#Created by Heather Musson on May 14, 2018 (Last updated May 15, 2018)
#References:
    #https://medium.freecodecamp.org/send-emails-using-code-4fcea9df63f

import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_EMAIL = 'email@email.com'
MY_PASSWORD = 'password'

#get contacts
def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding = 'utf-8')
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails

#import message
def read_template(filename):
    with open(filename, 'r', encodings = 'utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():
    #set up smtp
    s = smtplib.SMTP(host='your_host_address_here', port = 'port')
    s.starttls()
    s.login(MY_EMAIL, MY_PASSWORD)

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
        s.send_message(msg)
        del msg

    s.quit()
    
if __name__ == '__main__':
    main()
