# Custom-Email-Automation

Code that allows the user to send emails to a list of contacts (either in an excel file or a text file). The email can be customized for each recipient - currently set up to change who the email is addressed to but this can be modifed to add more customized fields. Some knowledge of Python and the Terminal is needed to modify and run this program.

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

You made add images using the following code, however change 'image1' to a unique identifer.

```
<img src="cid:image1">
```

2. Plain-text template
In messageplain.txt add the plain-text version of the message you wish to send. Again, wherever you want a custom field write the following code, however change 'PERSON_NAME' to a unique identifier. 

``` 
${PERSON_NAME}
```

In the plain-text message template no images or other unqiue HTML elements may be used.

3. Changes to emailauto.py
If you chose to add a custom field 'PERSON_NAME' must be changed to your unique identifer in the code below which can be found on lines 67 and 68 (shown below). If you chose to add more than one custom field additional .substitute must be called. If you chose no custom fields, lines 67 and 68 must be removed (either by deletion, or by making them comments - reccommended). In this example the text that we subsitute comes from a variable called 'name'  if you chose to change this in your code you must also change its name here.

```
messagehtml = message_template_html.substitute(PERSON_NAME = name.title())
messageplain = message_template_plain.substitute(PERSON_NAME = name.title())
```

4. More changes to emailauto.py
To set the subject of your email 'This is a test' on line 75 (shown below) must be changed.

```
msg['Subject'] = "This is a test"
```

5. Even more changes to emailauto.py
If you change the names of your variables on line 55 or 56 (show below) (dependent upon whether you choose to load your contacts from an excel file or a text file):

```
names, emails = get_contacts('mycontacts.txt')
names, emails = get_contacts_excel('contacts.xlsx', 'Sheet1')
```

you must also change their name on line 63 (shown below)

```
for name, email in zip(names, emails):
```

## Built With

* Python
