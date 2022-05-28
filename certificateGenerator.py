import pandas as pd
from pandas import ExcelWriter,ExcelFile
from PIL import Image,ImageDraw,ImageFont
import string
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import smtplib

fromEmail = "mail"
pwd = "pswd"
df_students = pd.read_excel(r'data.xlsx', sheet_name='Sheet1')
for i in df_students.index:
    certificate = Image.open('open.jpg')
    write = ImageDraw.Draw(certificate)
    color = 'rgb(0,0,0)'
    font = ImageFont.truetype('Sanchez-Regular.ttf', size=75)
    name = df_students['Name'][i]
    name.upper()
    write.text((500,720), name, fill=color, font=font)
    certificateName = "Certificate of "+ name + ".pdf"
    certificate.save(certificateName)
    toEmail = df_students['Email'][i]
    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = "Certificate Hackathon"

    body = '''Thank you ''' + name +''' participating in Open Meet

    Regards From
    OPEN'''

    msg.attach(MIMEText(body, 'plain'))
    attachmentName = name + ".pdf"
    with open(certificateName, "rb") as f:
            attach = MIMEApplication(f.read(),_subtype="pdf")
    attach.add_header('Content-Disposition','attachment',filename=attachmentName)
    msg.attach(attach)
    s = smtplib.SMTP('smtp.gmail.com', 587)   
    s.starttls()
    s.login(fromEmail, pwd)
    text = msg.as_string()
    s.sendmail(fromEmail, toEmail, text)

s.quit()