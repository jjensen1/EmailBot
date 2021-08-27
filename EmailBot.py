#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 11 14:57:32 2021
@author: junjensen
"""

# Send a multipart email with the appropriate MIME types.

import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

emailfrom = "SenderEmailName@email.com"
emailto = "RecipientEmailName@email.com, RecipientEmailName01@email.com"
Cc = 'RecipientEmailName02@email.com'
body = 'This is the body of the email.'
body = MIMEText(body) # convert the body to a MIME compatible string
fileToSend = "YourExcelFile.csv"
username = "SenderEmailName@email.com"
password = "SenderEmailPw"

msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg["Cc"] = Cc
msg["Subject"] = "Hi. Read this email and check out the attachment"
msg.attach(body) # attach it to your main message


# Below if/elif/else statement accounts for a range of file type attachments.

ctype, encoding = mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    # No guess could be made, or the file is encoded (compressed), so
    # use a generic bag-of-bits type.
    ctype = "application/octet-stream"

maintype, subtype = ctype.split("/", 1)

if maintype == "text":
    fp = open(fileToSend)
    # Note: we should handle calculating the charset
    attachment = MIMEText(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "image":
    fp = open(fileToSend, "rb")
    attachment = MIMEImage(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "audio":
    fp = open(fileToSend, "rb")
    attachment = MIMEAudio(fp.read(), _subtype=subtype)
    fp.close()
else:
    fp = open(fileToSend, "rb")
    attachment = MIMEBase(maintype, subtype)
    attachment.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition", "attachment", filename=fileToSend) #add report header with the file name
msg.attach(attachment)

server = smtplib.SMTP("smtp.gmail.com:587")
server.starttls()
server.login(username,password)
server.sendmail(emailfrom, emailto.split(',') + Cc.split(','), msg.as_string())
server.quit()
