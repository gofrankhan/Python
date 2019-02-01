import smtplib
from email.message import EmailMessage
import ssl

msg = EmailMessage()
msg.set_content("Hello this is a python mail sending test")

msg['Subject'] = 'here is subject'
msg['From'] = 'gofran0123@gmail.com'
msg['To'] = 'gofran3951@diu.edu.bd'

s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
s.login('gofran0123@gmail.com', password)
s.send_message(msg)
s.quit()
