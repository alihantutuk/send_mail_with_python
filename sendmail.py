import smtplib
from typing import Optional, List
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

def send_mail(email: str, password: str,send_to_email: str, message: str,subject: Optional[str]="Script Output")-> None:
    """
    This function sends e-mail according to the given appropriate parameters

    :param email: User e-mail address
    :param password: User e-mail password
    :param send_to_email: Destination e-mail address
    :param message: A text which you want to send
    :param subject: Mail description
    :return: None
    """

    server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to the server
    server.starttls()  # Use TLS
    server.login(email, password)  # Login to the email server

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = send_to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    server.send_message(msg)
    server.quit() # Logout of the email server


def send_multiple_mail(email: str, password: str,send_to_email_addresses: List[str], message: str,subject: Optional[str]="Script Output")-> None:
    """
    This function sends e-mails to multiple destination addresses according to the given appropriate parameters
    :param email: User e-mail address
    :param password: User e-mail password
    :param send_to_email_addresses: Destination e-mail addresses list
    :param message: A text which you want to send
    :param subject: Mail description
    :return: None
    """
    server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to the server
    server.starttls()  # Use TLS
    server.login(email, password)  # Login to the email server

    for mail in range(len(send_to_email_addresses)):
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = send_to_email_addresses[mail]
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))
        server.send_message(msg)

    server.quit()  # Logout of the email server

if __name__== '__main__':
    send_mail(os.getenv('EMAIL'),os.getenv('PASSWORD'),os.getenv('TARGET_EMAIL'),"Main","Subject")

