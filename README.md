# Custom-Email-Automation

Code that allows the user to send emails to a list of contacts (either in an excel file or a text file). The email can be customized for each recipient - currently set up to change who the email is addressed to but this can be modified to add more customized fields. Some knowledge of Python and the Terminal is needed to modify and run this program.

## What is this project?

This project was created while working at a company where I was required to send large amounts of marketing emails each day. With the desire to speed up a repetitive process this project was created. Currently the code is set up for the general case, but with some knowledge of Python and the information contained in this readme it can be modified for a specific case.

## Folder and File Structure
```
./
├── .gitignore
├── contacts.xlsx       * Place to store contacts if reading from an excel file.
├── image.jpg           * General placeholder image.
├── messagehtml.txt     * HTML version of the message to be sent.
├── messageplain.txt    * Plain-text version of the message to be sent.
├── mycontacts.txt      * Place to store contacts if reading from a text file.
├── README.md
└── references.txt      * File containing links to articles that were helpful in creating this project.
```

## Getting Started

These instructions will get you a copy of the project up and running on your local machine

### Dependencies

This project is built using Python Version 3.6.5 which can be installed from [here](https://www.python.org/downloads/). The python library, xlrd, must be installed as well, which can be done through the following command:

```
pip3 install xlrd
```

### Modifications

#### Setup your email server

Several modifications must be made to set up your email connection.

1. Host server name and port
You must must find out what the SMTP server name and port is for your email client. For example, Gmail's server name is smtp.gmail.com, and the secure port is 465. On line 48 of the file emailauto.py (shown below) change to default gmail settings to your required server and port.

```
s = smtplib.SMTP_SSL(host='smtp.gmail.com', port = '465')
```

2. Email username and password
You must change the default email and password on line 50 of the file emailauto.py (shown below) to that of your own email.

```
s.login('EMAIL', 'PASSWORD')
```

3. Change email in 'from' field
You must finally change the default email on line 73 of the file emailauto.py (shown below) to that of your own email:

```
msg['From'] = 'EMAIL
```

#### Customize your email template

1. HTML template
In messagehtml.txt add the html of your custom message. Wherever you want a custom field write the following code, however change 'PERSON_NAME' to a unique identifier.

```
${PERSON_NAME}
```

You made add images using the following code, however change 'image1' to a unique identifier.

```
<img src="cid:image1">
```

2. Plain-text template
In messageplain.txt add the plain-text version of the message you wish to send. Again, wherever you want a custom field write the following code, however change 'PERSON_NAME' to a unique identifier.

```
${PERSON_NAME}
```

In the plain-text message template no images or other unique HTML elements may be used.

3. Updating variables based on the number of custom fields

a) If you chose to have one custom fields
The general example is set up with only one custom field. If you chose to change the identifier of the variable used (currently called 'names'), see the Section 'Lines that must be changed' to see where you must also change its identifier.

b) If you chose to have no custom fields
If you chose to have no custom fields then all references to the variable entitled 'names' must be removed. See the Section 'Lines that must be changed' to see where must removed the variable.

c) If you chose to have more than one custom fields
Additional variables must be added for each additional custom field. For example if you have 3 custom fields then 2 additional variables must be added since the general example already includes 1 custom field variable. See the Section 'Lines that must be changed' to see where to add you additional variables.

##### Lines that must be changed
Line 12:
```
names = []
```

Line 16
```
names.append(a_contact.split()[0])
```

Line 22:
```
names = []
```

Line 33:
```
names.append(row[0])
```

Line 36:
```
return names, emails
```

Lines 55 and 56:
```
#names, emails = get_contacts('mycontacts.txt')
names, emails = get_contacts_excel('contacts.xlsx', 'Sheet1')
```

Line 63:
```
for name, email in zip(names, emails):
```

Lines 67 and 68:
```
messagehtml = message_template_html.substitute(PERSON_NAME = name.title())
messageplain = message_template_plain.substitute(PERSON_NAME = name.title())
```

4. Changing the subject of your email
To change the subject of your email, 'This is a test' on line 75 (shown below) must be changed.

```
msg['Subject'] = "This is a test"
```

6. Adding images
Currently the code supports one image in the HTML verison of your message. This can be changed to either remove the image, or add additional images.

a)

## Built With

* Python
