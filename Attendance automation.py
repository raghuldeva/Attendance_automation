import cv2
import pytesseract
import pandas as pd
import numpy as np
import pdfkit
import smtplib

from styleframe import StyleFrame
data = pd.read_excel('data.xlsx')

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

img = cv2.imread("C:\\Users\\Raghul Devaraj\\PycharmProjects\\Capture.png")
grey = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

ret, thresh1 = cv2.threshold(grey, 0, 255, cv2.THRESH_OTSU)


rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (13, 13))

dilation = cv2.dilate(thresh1, rect_kernel, iterations=1)

cv2.imshow("My img1", dilation)
cv2.waitKey(0)

contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL,
                                       cv2.CHAIN_APPROX_NONE)

im2 = img.copy()

Names = ['AISHWARYA' ,'AKSHAYA' ,'ANJANA' ,'ARASU' ,'BALA MURUGAN' ,'BHARANIPRABHA' ,'BRUNTHA' ,'CAROLYNSUSANALEX' ,'DHIVYASRI' ,'DIVYA' ,'ELAKKIYA' ,'GOVINDARAJ' ,'GOWTHAM' ,'GURUSHIYA' ,'JAYAPRAKASH' ,'KEERTHI VAASAN' ,'KOKILA' ,'MANICKAVASAGAM' ,'MANIRAJAN' ,'IBRAHIM' ,'NARMADHA' ,'RAGAVI' ,'RAGHUL' ,'RAJAGOPAL' ,'RANJITH' ,'RAVISANKAR' ,'RIHA' ,'SAMYUKTHA' ,'SATHIS KUMAR' ,'SAVITHA' ,'SENTHIL KUMAR' ,'SRI HARINI' ,'SUDHARSAN' ,'SWATHI' ,'UDHAYAMARISH' ,'VARSHINI' ,'VIGNESH','VISHAL']
present = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    rect = cv2.rectangle(im2, (x, y), (x + w, y + h), (0, 255, 0), 2)
    cropped = im2[y:y + h, x:x + w]
    file = open("recognized.txt", "a")
    text = pytesseract.image_to_string(cropped)
    text = ''.join([i for i in text if i.isalpha() or i.isdigit() or i == " "])
    text = text.lstrip()
    text = text.upper()
    if text != "":
        present.append(text)

print(present)

present_names = []
for i in present:
    if '21CSEG' in i:
        roll_index = i.index('21CSEG')
        reg_no = i[roll_index:roll_index + 8]
        i = i.replace(reg_no, "")
        present_names.append(i.lstrip())
    else:
        present_names.append(i.lstrip())

print(present_names)

status = []

for i in Names:
    for j in present_names:
        if i in j:
            status.append("Present")
            break
    else:
        status.append("Absent")

print(status)

data['Status'] = np.array(status)
print(data)


def row_style(row):
    if row.Status == 'Present':
        return pd.Series('background-color: green', row.index)
    else:
        return pd.Series('background-color: red', row.index)


data = data.style.apply(row_style, axis=1)

f = open('exp.html','w')
a = data.to_html()
f.write(a)
f.close()

pdfkit.from_file('exp.html', 'final.pdf')

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

body = '''
Respected Mam, 
    Please find the attachment containing attendance
'''

sender = '****************@gmail.com'
password = '*************'
receiver = '****************@gmail.com'

#Setup the MIME
message = MIMEMultipart()
message['From'] = sender
message['To'] = receiver

message['Subject'] = 'Attendance : Principles of data science'
message.attach(MIMEText(body, 'plain'))

pdfname = 'final.pdf'


# open the file in bynary
binary_pdf = open(pdfname, 'rb')

payload = MIMEBase('application', 'octate-stream', Name=pdfname)
payload.set_payload((binary_pdf).read())

# enconding the binary into base64
encoders.encode_base64(payload)

# add header with pdf name
payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
message.attach(payload)

#use gmail with port
session = smtplib.SMTP('smtp.gmail.com', 587)

#enable security
session.starttls()

#login with mail_id and password
session.login(sender, password)

text = message.as_string()
session.sendmail(sender, receiver, text)
session.quit()
print('Mail Sent')

