import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from googleapiclient.discovery import build

def send_mail_html(creds, sender, to, subject, html_body, display_name):
    service = build('gmail', 'v1', credentials=creds)
    message = MIMEMultipart("alternative")
    message['To'] = to
    message['From'] = formataddr((display_name, sender))
    message['Subject'] = subject

    msg_html = MIMEText(html_body, "html")
    message.attach(msg_html)
    
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId="me", body={'raw': raw}).execute()
