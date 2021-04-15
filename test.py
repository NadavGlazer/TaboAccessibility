from main import pdf_to_txt
from email.message import EmailMessage
import mimetypes
import smtplib
import threading


class Threads(threading.Thread):
        def __init__(self, pdf_file_name):
            threading.Thread.__init__(self)
            self.pdf_file_name = pdf_file_name

        def run(self):
            pdf_to_txt(self.pdf_file_name)


#call this after you got the pdf file
pdf_file = "s.pdf"
vars()['Thread' +pdf_file[:-4]] = Threads(pdf_file)
vars()['Thread' +pdf_file[:-4]].Start()
#after the excel file is ready, call email-send function

def email_send(excel_file_name):
    message = EmailMessage()
    sender_email = "your-name@gmail.com"
    sender_password = "password"
    recipient = "example@example.com"
    message['From'] = sender_email
    message['To'] = recipient

    message['Subject'] = excel_file_name[:-5] + " results"
    mime_type, _ = mimetypes.guess_type(excel_file_name)
    mime_type, mime_subtype = mime_type.split('/')

    with open(excel_file_name, 'rb') as file:
        message.add_attachment(file.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename='something.pdf')

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.set_debuglevel(1)
    mail_server.login(sender_email, sender_password)
    mail_server.send_message(message)
    mail_server.quit()


