import streamlit as st
import json
import time
import random
import pandas as pd
import re
from datetime import datetime
import base64

# --- NEW IMPORTS FOR ROBUST HTML EMAILS ---
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.message import EmailMessage

# Google Libraries
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

st.set_page_config(page_title="Simple Merge", page_icon="👔", layout="wide")

# --- 1. CORE LOGIC ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/drive.readonly', 
    'https://www.googleapis.com/auth/spreadsheets'
]

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

def get_jd_html(creds, doc_id):
    drive_service = build('drive', 'v3', credentials=creds)
    html_content = drive_service.files().export(fileId=doc_id, mimeType='text/html').execute()
    decoded_html = html_content.decode('utf-8')
    
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
    
def get_full_sheet_data(creds, sheet_id, sheet_name):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        return res.get('values', []) if res.get('values') else []
    except: return []

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

                thread_messages.append({
                    "From": sender,
                    "Date": headers.get('Date', ''),
                    "Body": "".join(body_html) if body_html else "No HTML Body",
                    "Attachments": attachments,
                    "Message_ID": msg['id']
                })

            first_headers = {h['name']: h['value'] for h in messages[0]['payload']['headers']}
            last_headers = {h['name']: h['value'] for h in messages[-1]['payload']['headers']}
            
            master_inbox.append({
                "Account": email,
                "Thread_ID": th['id'],
                "Subject": first_headers.get('Subject', 'No Subject'),
                "Vendor_Email": vendor_email,
                "Messages": thread_messages,
                "Last_RFC_Message_ID": last_headers.get('Message-ID', '')
            })
            
    return master_inbox

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

# --- 2. UI SETUP ---
st.title("👔 Simple Merge")

if "code" in st.query_params:
    code = st.query_params["code"]
    email_trying = st.query_params.get("state", "Unknown Account")
    try:
        redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
        flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
        flow.fetch_token(code=code)
        
        st.success(f"✅ LOGIN SUCCESS FOR: {email_trying}")
        st.warning("⬇️ COPY THIS TOKEN BELOW AND PASTE INTO SECRETS ⬇️")
        st.code(flow.credentials.to_json(), language="json")
        st.stop()
    except Exception as e:
        st.error(f"Login Error: {str(e)}")

# --- 3. TABS ---
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False
tab_run, tab_preview, tab_auth, tab_inbox = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts", "📥 Inbox"])

# --- TAB: ACCOUNTS ---
with tab_auth:
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        if col2.button("Login / Refresh", key=f"login_{email}"):
            redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
            flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
            url, _ = flow.authorization_url(prompt='consent', state=email)
            st.link_button("👉 Start Auth", url)

# --- TAB: PREVIEW ---
with tab_preview:
    admin_email = json.loads(st.secrets["DUMMY_ACCOUNTS"])[0]
    creds = load_creds(admin_email)
    if creds:
        try:
            subj, html_body = get_jd_html(creds, st.secrets["DOC_ID"])
            st.info(f"📄 Template: {subj}")
            st.markdown("**Personalized Preview (with HTML formatting):**")
            preview_html = html_body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
            st.html(preview_html)
        except Exception as e: st.error(f"Could not load preview: {e}")
    else: st.warning("Connect your first account to preview.")

# --- TAB: OPERATIONS ---
with tab_run:
    all_acc = json.loads(st.secrets["DUMMY_ACCOUNTS"])
    
    with st.sidebar:
        st.header("⚙️ Configuration")
        display_name = st.text_input("Send As Name", value=st.secrets.get("DISPLAY_NAME", "Recruitment Team"))
        sel_acc = st.multiselect("Active Senders", all_acc, default=all_acc)
        limit = st.number_input("Max Per Account", 1, 500, 20)
        delay = st.number_input("Round Delay (s)", 5, 600, 20)
        is_dry = st.toggle("🧪 Dry Run (Safe)", value=True)

    col_btn, col_stop = st.columns([1, 4])
    start = col_btn.button("🔥 LAUNCH", type="primary", use_container_width=True)
    if col_stop.button("🛑 STOP CAMPAIGN", type="secondary"): st.session_state.stop_clicked = True

    if start:
        st.session_state.stop_clicked = False
        active_data = []
        with st.status("🔍 Checking System...") as status:
            admin_creds = load_creds(all_acc[0])
            subj, body_tmpl = get_jd_html(admin_creds, st.secrets["DOC_ID"])
            for s in sel_acc:
                c = load_creds(s)
                if c:
                    rows = get_full_sheet_data(c, st.secrets["SHEET_ID"], f"filter{all_acc.index(s)}")
                    active_data.append({"email": s, "creds": c, "rows": rows, "idx": 0})
            status.update(label="System Ready!", state="complete")

        dashboard_df = pd.DataFrame([{"Account": s["email"], "Target": "-", "Sent": 0, "Status": "Ready"} for s in active_data]).set_index("Account")
        prog_bar = st.progress(0, text="Progress")
        table_ui = st.empty()
        
        total_goal = sum([min(len(s["rows"]), limit) for s in active_data])
        sent_total = 0

        for r_num in range(limit):
            if st.session_state.stop_clicked: break
            
            round_active = False
            for s in active_data:
                if s["idx"] < len(s["rows"]):
                    round_active = True
                    row = s["rows"][s["idx"]]
                    target = row[0]
                    comp = row[1] if len(row) > 1 else "Your Company"
                    role = row[2] if len(row) > 2 else "the open position"
                    fname = target.split('@')[0].split('.')[0].capitalize()
                    
                    final_body = body_tmpl.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                    
                    dashboard_df.at[s["email"], "Target"] = target
                    dashboard_df.at[s["email"], "Status"] = "📨 Sending..."
                    table_ui.dataframe(dashboard_df, use_container_width=True)

                    try:
                        if not is_dry: 
                            send_mail_html(s["creds"], s["email"], target, subj, final_body, display_name)
                        
                        s["idx"] += 1
                        sent_total += 1
                        dashboard_df.at[s["email"], "Sent"] = s["idx"]
                        dashboard_df.at[s["email"], "Status"] = "✅ Sent"
                    except Exception as e:
                        error_msg = str(e).split(']')[0] 
                        dashboard_df.at[s["email"], "Status"] = f"❌ {error_msg}"
                        st.error(f"Detailed Error for {s['email']}: {e}")
                    
                    time.sleep(1)
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    prog_bar.progress(sent_total/total_goal, text=f"Sent {sent_total} / {total_goal}")

            if not round_active: break
            
            if r_num < limit - 1:
                human_delay = delay + random.randint(-2, 2)
                for sec in range(human_delay, 0, -1):
                    if st.session_state.stop_clicked: break
                    for s in active_data: 
                        if "Auth" not in str(dashboard_df.at[s["email"], "Status"]) and "Error" not in str(dashboard_df.at[s["email"], "Status"]):
                            dashboard_df.at[s["email"], "Status"] = f"⏳ {sec}s"
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    time.sleep(1)

        st.balloons()

# --- TAB: INBOX ---
with tab_inbox:
    col_head1, col_head2 = st.columns([4, 1])
    col_head1.subheader("📥 Unified Vendor Inbox")
    if col_head2.button("🔄 Refresh", use_container_width=True):
        fetch_all_threads.clear()
        st.rerun()
        
    emails = fetch_all_threads()
    
    if not emails:
        st.info("No active vendor conversations found.")
    else:
        # initialize session state to handle active conversation selections
        if "selected_inbox_idx" not in st.session_state:
            st.session_state.selected_inbox_idx = 0
            
        # safety catch if the inbox list shrinks after a refresh
        if st.session_state.selected_inbox_idx >= len(emails):
            st.session_state.selected_inbox_idx = 0

        # expanded left column for multi-line text
        col_list, col_view = st.columns([1.2, 2.5])
        
        with col_list:
            st.markdown("##### Conversations")
            # scrollable container so the whole page doesn't stretch down
            with st.container(height=500, border=False):
                for i, em in enumerate(emails):
                    sender = em['Vendor_Email'].split('<')[0].strip()[:25]
                    subj = em['Subject'][:32]
                    subj = subj + "..." if len(em['Subject']) > 32 else subj
                    acc = em['Account'].split('@')[0]
                    
                    # strip html tags from the body of the last message to create a clean snippet
                    last_msg_raw = em['Messages'][-1]['Body']
                    clean_snippet = re.sub('<[^<]+>', '', last_msg_raw).strip()
                    snippet = clean_snippet[:40] + "..." if len(clean_snippet) > 40 else clean_snippet
                    
                    # \n renders natively as multi-line inside Streamlit buttons
                    label = f"👤 {sender}\n📄 {subj}\n💬 {snippet}\n📥 {acc}"
                    
                    btn_type = "primary" if st.session_state.selected_inbox_idx == i else "secondary"
                    
                    if st.button(label, key=f"btn_thread_{em['Thread_ID']}", use_container_width=True, type=btn_type):
                        st.session_state.selected_inbox_idx = i
                        st.rerun()
            
        selected_thread = emails[st.session_state.selected_inbox_idx]
        
        with col_view:
            st.markdown(f"### {selected_thread['Subject']}")
            st.caption(f"**Vendor:** `{selected_thread['Vendor_Email']}` | **Via:** `{selected_thread['Account']}`")
            st.divider()
            
            with st.container(height=500, border=True):
                for msg in selected_thread["Messages"]:
                    is_me = selected_thread["Account"] in msg["From"]
                    
                    with st.container(border=True):
                        if is_me:
                            st.markdown(f"🟢 **You** (`{msg['Date']}`)")
                        else:
                            st.markdown(f"🔵 **Vendor** - {msg['From']} (`{msg['Date']}`)")
                        
                        st.html(msg["Body"])
                        
                        if msg["Attachments"]:
                            for att in msg["Attachments"]:
                                st.download_button(label=f"📎 {att['filename']}", data=att['data'], file_name=att['filename'], key=f"att_{msg['Message_ID']}")

            st.divider()
            st.markdown("#### Quick Reply")
            reply_body = st.text_area("Message:", key=f"reply_{selected_thread['Thread_ID']}")
            
            if st.button("Send Reply", type="primary", key=f"btn_{selected_thread['Thread_ID']}"):
                if reply_body.strip():
                    with st.spinner("Sending..."):
                        send_inbox_reply(
                            email=selected_thread["Account"],
                            thread_id=selected_thread["Thread_ID"],
                            rfc_message_id=selected_thread["Last_RFC_Message_ID"],
                            to_address=selected_thread["Vendor_Email"],
                            subject=selected_thread["Subject"],
                            body_text=reply_body
                        )
                    st.success("Reply sent and attached to thread!")
                    time.sleep(1)
                    fetch_all_threads.clear()
                    st.rerun()
                else:
                    st.error("Cannot send an empty message.")
