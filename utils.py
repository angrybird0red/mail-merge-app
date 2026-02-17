import streamlit as st
import json
import base64
import re
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.message import EmailMessage

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# --- CONSTANTS ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/drive.readonly', 
    'https://www.googleapis.com/auth/spreadsheets'
]

# --- AUTHENTICATION ---
def get_client_config():
    return json.loads(st.secrets["gcp_service_account"])

def load_creds(email):
    safe_email = email.replace("@", "_").replace(".", "_").upper()
    secret_key = f"TOKEN_{safe_email}"
    
    if secret_key in st.secrets:
        try:
            token_data = st.secrets[secret_key]
            try:
                token_info = json.loads(token_data)
            except:
                token_info = token_data
           
            if isinstance(token_info, str):
                token_info = json.loads(token_info)

            if isinstance(token_info, dict) and "client_id" not in token_info:
                main_config = json.loads(st.secrets["gcp_service_account"])
                app_info = main_config.get("web", main_config.get("installed", {}))
                token_info["client_id"] = app_info.get("client_id")
                token_info["client_secret"] = app_info.get("client_secret")
            
            creds = Credentials.from_authorized_user_info(token_info)
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
            return creds
        except Exception as e:
            st.error(f"⚠️ Token Error for {email}: {e}")
            return None
    return None

# --- GOOGLE DOCS & HTML PARSING ---
def get_jd_html(creds, doc_id):
    drive_service = build('drive', 'v3', credentials=creds)
    html_content = drive_service.files().export(fileId=doc_id, mimeType='text/html').execute()
    decoded_html = html_content.decode('utf-8')
    
    # Cleaning regex
    decoded_html = re.sub(r'margin-top:\s*[\d\.]+(pt|px|cm|in);?', 'margin-top: 0 !important;', decoded_html)
    decoded_html = re.sub(r'margin-bottom:\s*[\d\.]+(pt|px|cm|in);?', 'margin-bottom: 0 !important;', decoded_html)
    decoded_html = re.sub(r'padding-top:\s*[\d\.]+(pt|px|cm|in);?', 'padding-top: 0 !important;', decoded_html)
    decoded_html = re.sub(r'padding-bottom:\s*[\d\.]+(pt|px|cm|in);?', 'padding-bottom: 0 !important;', decoded_html)
    decoded_html = re.sub(r'\.c\d+\s*{[^}]+}', '', decoded_html)
    decoded_html = re.sub(r'(<(h[1-6]|p)[^>]*>)', r'\1', decoded_html, count=1)
    
    if "<body" in decoded_html:
        decoded_html = re.sub(r'<body[^>]*>', '<body style="margin:0; padding:0; background-color:#ffffff;">', decoded_html)

    style_fix = """
    <style>
        body { margin: 0 !important; padding: 0 !important; }
        body > *, body > div > * { margin-top: 0 !important; padding-top: 0 !important; }
        body, td, p, h1, h2, h3 { font-family: Arial, Helvetica, sans-serif !important; color: #000000 !important; }
        p { margin-bottom: 8px !important; margin-top: 0 !important; }
        ul, ol { margin-top: 0 !important; margin-bottom: 8px !important; padding-left: 25px !important; }
        li { margin-bottom: 2px !important; }
        li p { display: inline !important; margin: 0 !important; }
        h1, h2, h3 { margin-bottom: 10px !important; margin-top: 15px !important; }
        h1:first-child, h2:first-child, h3:first-child { margin-top: 0 !important; }
    </style>
    """
    
    clean_html = style_fix + decoded_html
    docs_service = build('docs', 'v1', credentials=creds)
    doc = docs_service.documents().get(documentId=doc_id).execute()
    return doc.get('title'), clean_html

# --- GOOGLE SHEETS ---
def get_full_sheet_data(creds, sheet_id, sheet_name):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        return res.get('values', []) if res.get('values') else []
    except: return []

def get_send_log(creds, sheet_id):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range="SendLog!A:B").execute()
        values = res.get('values', [])
        return set((row[0].strip(), row[1].strip()) for row in values if len(row) >= 2)
    except Exception as e:
        print(f"Error reading SendLog: {e}")
        return set()

def append_to_send_log(creds, sheet_id, target_email, subject, sender_account):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        timestamp = datetime.now(ist_tz).strftime('%Y-%m-%d %I:%M:%S %p IST')
        
        body = {'values': [[target_email, subject, sender_account, timestamp]]}
        sheets.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range="SendLog!A:D",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()
    except Exception as e:
        print(f"Error appending to SendLog: {e}")

# --- GMAIL SENDING & READING ---
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
