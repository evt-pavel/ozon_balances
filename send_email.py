import smtplib
from setting import sender, password, recipient
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import os



def send_email(filename):

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = f'Остатки ozon {date.today().strftime("%d-%b-%Y")}'

        attachment = open(filename, 'rb')
        xlsx = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        xlsx.set_payload(attachment.read())
        encoders.encode_base64(xlsx)
        xlsx.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(xlsx)

        server.sendmail(sender, recipient, msg.as_string())
        os.remove(filename)
        return 'The message was sand successfully'

    except Exception as ex:
        return f'Error {ex}'





