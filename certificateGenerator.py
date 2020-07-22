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

# the email address and password of sender
fromEmail = "Sender's Email id here"
pwd = "Sender's Email password here"


# opening the data about students in dataframe using pandas
df_students = pd.read_excel(r'data.xlsx', sheet_name='Sheet1')

for i in df_students.index:
    # getting the certificate template
    certificate = Image.open('demo.jpeg')
    # using the image draw fuction from PIL to write the name of student
    write = ImageDraw.Draw(certificate)
    color = 'rgb(0,0,0)'
    # setting the fon type and size using image font from PIL
    font = ImageFont.truetype('Raleway-Regular.ttf', size=50)
    # getting the name of each student from dataframe
    name = df_students['Name'][i]
    name.upper()
    # putting the x,y coordinates from where the name is to be started
    write.text((680,515), name, fill=color, font=font)
    # saving the file as pdf
    certificateName = "Certificate of "+ name + ".pdf"
    certificate.save(certificateName)

    # getting the email address of the student from dataframe
    toEmail = df_students['Email'][i]
    # MIMEMultipart is used to fill various fields required for mailing 
    msg = MIMEMultipart()

    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = "Certificate for XYZ Event"

    # making a string to store body of the mail to be sent
    body = '''Thank you ''' + name +''' participating in XYZ Event

    Regards From
    College Name/Event Name'''

    # attaching the body with msg
    msg.attach(MIMEText(body, 'plain'))

    # attaching the certificate to be sent
    attachmentName = name + ".pdf"
    with open(certificateName, "rb") as f:
            attach = MIMEApplication(f.read(),_subtype="pdf")
    attach.add_header('Content-Disposition','attachment',filename=attachmentName)
    msg.attach(attach)

    # attachment = open(certificateName, errors='ignore')
    # p = MIMEBase('application', 'octet-stream')

    # # changing payload in encoded form  
    # p.set_payload((attachment).read())
    # # encoding into base64
    # encoders.encode_base64(p)

    # p.add_header('Content-Disposition', "attachment; filename = %s" %attachmentName)

    # # attaching payload with msg
    # msg.attach(p)

    # creating SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
    # starting TLS to have security to send mail
    s.starttls()

    # Logging in
    s.login(fromEmail, pwd)
    # converting the Multipart message into string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromEmail, toEmail, text)

# terminating the session
s.quit()