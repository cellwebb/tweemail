import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib


class Email:
    '''
    Parameters
    ----------
    sender : str
        address from which to send email
    recipients : str or list
        addresses of email recipients
    subject : str
        email subject
    body : str
        email body
    attachments : str or list
        filepaths of email attachments
    host : str
        SMTP host address
    port : int
        SMTP host port
    body_type : str, optional
        accepted options are 'html' and 'plain'
    '''
    def __init__(self, sender=None, recipients=None, subject=None, body=None,
                 attachments=None, host=None, port=None, body_type='html'):
        self.msg = MIMEMultipart()
        if sender:
            self.set_sender(sender)
        if recipients:
            self.set_recipients(recipients)
        if subject:
            self.set_subject(subject)
        if body:
            self.attach_body(body, body_type)
        if attachments:
            self.attach_files(attachments)
        if host:
            self.set_host(host)
        if port:
            self.set_port(port)

    def attach_body(self, body, body_type='html'):
        '''
        Attaches email body

        body : str
            email body
        body_type : str, optional
            accepted options are 'html' and 'plain'
        '''
        self.msg.attach(MIMEText(body, body_type))

    def attach_files(self, attachments):
        '''
        Attaches file(s) to email

        attachments : str or list
            filepaths of email attachments
        '''
        if isinstance(attachments, str):
            attachments = [attachments]

        for attachment in attachments:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(open(attachment, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=os.path.split(attachment)[1])
            self.msg.attach(part)

    def send(self):
        '''
        Sends email to recipients
        '''
        assert (self.msg['To'] and self.msg['From']), (
            'A sender and at least one recipient are required.')
        assert (hasattr(self, 'host')), (
            'You must set your host connection string before sending')

        if not hasattr(self, 'port'):
            print('Port not set, using default value of 0')
            self.port = 0

        s = smtplib.SMTP(self.host, self.port)
        try:
            s.sendmail(self.msg['From'], self.msg['To'], self.msg.as_string())
            print(f"Email sent to: {self.msg['To']}")
        except Exception as e:
            print('An error occured and email was not sent')
            print(e)
        finally:
            s.quit()

    def set_host(self, host):
        '''
        Sets host address for SMTP connection

        host : str
            SMTP host address
        '''
        self.host = host

    def set_port(self, port):
        '''
        Sets host port for SMTP connection

        port : int
            SMTP host port
        '''
        self.port = port

    def set_recipients(self, recipients):
        '''
        Sets email recipient(s)

        recipients : str or list
            addresses of email recipients
        '''
        if isinstance(recipients, str):
            recipients = [recipients]

        self.msg['To'] = ', '.join(recipients)

    def set_sender(self, sender):
        '''
        Sets email sender

        sender : str
            address from which to send email
        '''
        self.msg['From'] = sender

    def set_subject(self, subject):
        '''
        Sets email subject

        subject : str
            email subject
        '''
        self.msg['Subject'] = subject
