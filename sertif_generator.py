import os
from os import listdir
from os.path import isfile, join

import pandas as pd

from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

from natsort import natsorted

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

cwd = os.getcwd()

template = 'template.png'
spreadsheet_file = 'datalist.xlsx'
name_font = ImageFont.truetype('RobotoMono.ttf', 150)  # font name and size
font_color = (71, 48, 149)  # RGB value
output_folder_name = 'Sertif Files'

img_template = Image.open(template)
img_w, img_h = img_template.size

# SMTP server for gmail.
server = smtplib.SMTP(host='smtp.gmail.com', port=587)

sender_name = 'Sender Name'
subject = 'Mail Subject'

body = ("""
<html>
    <head></head>
    <body>
        <p>Hi!<br />
          How are you?<br />
          Here is the <a href="http://www.gitlab.com/deutan/">link</a> you wanted.
        </p>
    </body>
</html>
""")

spreadsheet_path = os.path.abspath(spreadsheet_file)
names_data = pd.read_excel(spreadsheet_path)

# list all values from "Nama" column
name_list = names_data["Nama"].values
names_count = len(name_list)

output_folder_path = os.path.join(cwd, output_folder_name)

def createOutputFolder():
    if not os.path.exists(output_folder_path):
        os.mkdir(output_folder_path)


def insertNames():
    createOutputFolder()
    n = 0
    pdf_paths = []
    for i in name_list:
        img_template = Image.open(template)
        draw = ImageDraw.Draw(img_template)

        name = name_list[n]
        # Change the value of x and y to adjust the text position
        name_w, name_h = name_font.getsize(name)
        x = (img_w - name_w)/2  # Center horizontally
        y = 970
        name_pos = (x, y)

        draw.text(name_pos, name.upper(), font_color, font = name_font)
        output_filename = str(n) + str(name) + ".pdf"
        pdf_path = os.path.join(cwd, output_folder_path, output_filename)
        img_template.save(pdf_path, "PDF")
        
        pdf_paths.append(pdf_path)
        print("Output:", output_filename)
        n += 1


def sendEmails():
    sender_mail = input('Sender Email: ')
    sender_pass = input('Password: ')
    server.starttls()
    server.login(sender_mail, sender_pass)
    pdf_paths = natsorted([os.path.join(output_folder_path, fn) for fn in next(os.walk(output_folder_path))[2]])
    pdf_names = natsorted([f for f in listdir(output_folder_path) if isfile(join(output_folder_path, f))])

    n = 0
    for index, row in names_data.iterrows():
        filename = pdf_names[n]
        msg = MIMEMultipart()
        msg['From'] = sender_name
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        toaddr = row["Email"]
        attachment = open(pdf_paths[n], "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)
        text = msg.as_string()
        server.sendmail(sender_name, toaddr, text)

        print (toaddr + ": " + filename)
        n += 1


def runScript():
    a = input("""
    Menu:
    1 : Generate certificates
    2 : Send mails
    Input your choice [1/2]: """)
    if a == '1':
        insertNames()
        b = input('Continue to mail generated certificates? [y/n]')
        if b.lower() == 'y':
            sendEmails()
        else:
            quit()
    if a == '2':
        sendEmails()
    else:
        quit()
    
# insertNames()
# sendEmails()
runScript()

server.quit()
