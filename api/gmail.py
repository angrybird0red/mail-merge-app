import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.message import EmailMessage
from googleapiclient.discovery import build
import streamlit as st
import json
from .auth import load_creds # Relative import from our own folder

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

def send_inbox_reply(email, thread_id, rfc_message_id, to_address, subject, body_text):
    creds = load_creds(email)
    service = build('gmail', 'v1', credentials=creds)
    
    message = EmailMessage()
    message.set_content(body_text)
    message['To'] = to_address
    message['Subject'] = subject if subject.startswith("Re:") else f"Re: {subject}"
    message['In-Reply-To'] = rfc_message_id
    message['References'] = rfc_message_id
    
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    service.users().messages().send(userId='me', body={'raw': raw_message, 'threadId': thread_id}).execute()
    service.users().messages().modify(userId='me', id=rfc_message_id, body={'removeLabelIds': ['UNREAD']}).execute()

def parse_email_parts(service, user_id, msg_id, parts, attachments, body_html):
    for part in parts:
        mime_type = part.get('mimeType')
        filename = part.get('filename')
        
        if mime_type == 'text/html' and not filename:
            data = part['body'].get('data')
            if data:
                decoded = base64.urlsafe_b64decode(data.encode('UTF-8')).decode('utf-8')
                body_html.append(decoded)
                
        if filename:
            attachment_id = part['body'].get('attachmentId')
            data = part['body'].get('data')
            if attachment_id:
                att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                data = att.get('data')
            if data:
                attachments.append({"filename": filename, "data": base64.urlsafe_b64decode(data.encode('UTF-8'))})
                
        if 'parts' in part:
            parse_email_parts(service, user_id, msg_id, part['parts'], attachments, body_html)

@st.cache_data(ttl=60)
def fetch_all_threads(max_threads_per_account=10):
    master_inbox = []
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    
    for email in accounts:
        creds = load_creds(email)
        if not creds: continue
        
        service = build('gmail', 'v1', credentials=creds)
        
        results = service.users().threads().list(userId='me', maxResults=max_threads_per_account, q='in:inbox -from:mailer-daemon').execute()
        threads = results.get('threads', [])
        
        for th in threads:
            th_data = service.users().threads().get(userId='me', id=th['id'], format='full').execute()
            messages = th_data.get('messages', [])
            if not messages: continue
            
            thread_messages = []
            vendor_email = "Unknown"
            
            for msg in messages:
                payload = msg['payload']
                headers = {h['name']: h['value'] for h in payload['headers']}
                
                sender = headers.get('From', 'Unknown')
                if email not in sender:
                    vendor_email = sender
                    
                attachments = []
                body_html = []
                
                if 'parts' in payload:
                    parse_email_parts(service, 'me', msg['id'], payload['parts'], attachments, body_html)
                else:
                    data = payload['body'].get('data')
                    if data: body_html.append(base64.urlsafe_b64decode(data.encode('UTF-8')).decode('utf-8'))

                msg_time = msg.get('internalDate', '0')

                thread_messages.append({
                    "From": sender,
                    "Date": headers.get('Date', ''),
                    "Internal_Date": msg_time,
                    "Body": "".join(body_html) if body_html else "No HTML Body",
                    "Attachments": attachments,
                    "Message_ID": msg['id'],
                    "Snippet": msg.get('snippet', '')
                })

            first_headers = {h['name']: h['value'] for h in messages[0]['payload']['headers']}
            last_headers = {h['name']: h['value'] for h in messages[-1]['payload']['headers']}
            last_msg_time = messages[-1].get('internalDate', '0')
            
            master_inbox.append({
                "Account": email,
                "Thread_ID": th['id'],
                "Subject": first_headers.get('Subject', 'No Subject'),
                "Vendor_Email": vendor_email,
                "Messages": thread_messages,
                "Last_Message_Time": last_msg_time,
                "Last_RFC_Message_ID": last_headers.get('Message-ID', '')
            })
            
    master_inbox.sort(key=lambda x: int(x['Last_Message_Time']), reverse=True)
    return master_inbox
