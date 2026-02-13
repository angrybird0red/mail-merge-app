import streamlit as st
import json
import time
import pandas as pd
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
import base64

# Google Libraries
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

st.set_page_config(page_title="Mail Merge App", page_icon="📨")

# --- 1. SETUP & AUTH ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

def get_client_config():
    # We get this from the "Safe Box" (Secrets)
    if "gcp_service_account" in st.secrets:
        return json.loads(st.secrets["gcp_service_account"])
    else:
        st.error("🚨 Missing 'gcp_service_account' in Secrets!")
        st.stop()

def load_creds(email):
    # We look for the token in the "Safe Box" (Secrets)
    # The key format is TOKEN_EMAIL_ADDRESS (e.g., TOKEN_BOB_GMAIL_COM)
    safe_email = email.replace("@", "_").replace(".", "_").upper()
    secret_key = f"TOKEN_{safe_email}"
    
    if secret_key in st.secrets:
        token_info = json.loads(st.secrets[secret_key])
        creds = Credentials.from_authorized_user_info(token_info)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    return None

# --- 2. GOOGLE HELPERS ---
def get_jd(creds, doc_id):
    docs = build('docs', 'v1', credentials=creds)
    doc = docs.documents().get(documentId=doc_id).execute()
    title = doc.get('title')
    content = doc.get('body', {}).get('content', [])
    text = []
    for elem in content:
        if 'paragraph' in elem:
            for e in elem['paragraph']['elements']:
                if 'textRun' in e:
                    text.append(e['textRun'].get('content', ''))
    return title, "".join(text)

def get_sheet_col(creds, sheet_id, sheet_name):
    sheets = build('sheets', 'v4', credentials=creds)
    res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:A").execute()
    return [row[0] for row in res.get('values', []) if row]

def log_send(creds, sheet_id, row):
    sheets = build('sheets', 'v4', credentials=creds)
    sheets.spreadsheets().values().append(
        spreadsheetId=sheet_id, range="SendLog!A:D", 
        valueInputOption="RAW", body={"values": [row]}
    ).execute()

def send_mail(creds, sender, to, subject, body, display_name):
    service = build('gmail', 'v1', credentials=creds)
    msg = EmailMessage()
    msg.set_content(body)
    msg['To'] = to
    msg['From'] = formataddr((display_name, sender))
    msg['Subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={'raw': raw}).execute()

# --- 3. THE PHONE UI ---
st.title("📨 Mail Merge")

# TABS
tab1, tab2 = st.tabs(["🚀 Run", "⚙️ Setup Accounts"])

with tab2:
    st.write("### Account Manager")
    # Get list of accounts from Secrets
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    
    if not accounts:
        st.warning("⚠️ No accounts found! Add 'DUMMY_ACCOUNTS' to Secrets.")
    
    # Handle the "I just logged in" return
    if "code" in st.query_params:
        code = st.query_params["code"]
        email_trying = st.query_params.get("state")
        try:
            flow = Flow.from_client_config(
                get_client_config(), SCOPES, 
                redirect_uri="https://lalith-mail-merge.streamlit.app"
            )
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.success(f"✅ SUCCESS! Logged in as {email_trying}")
            st.write("👇 **COPY THIS CODE** and paste it into your Secrets file:")
            st.code(creds.to_json(), language="json")
            st.info(f"Save it as: `TOKEN_{email_trying.replace('@','_').replace('.','_').upper()}`")
        except Exception as e:
            st.error(str(e))

    # Show Login Buttons
    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Not Connected"
        col1.write(f"**{email}** - {status}")
        
        if not creds:
            if col2.button("Login", key=email):
                flow = Flow.from_client_config(
                    get_client_config(), SCOPES, 
                    redirect_uri="https://lalith-mail-merge.streamlit.app"
                )
                url, _ = flow.authorization_url(prompt='consent', state=email)
                st.link_button("Go to Google", url)

with tab1:
    st.write("### Start Sending")
    limit = st.slider("Emails per account", 1, 50, 20)
    delay = st.number_input("Delay (seconds)", 5, 60, 20)
    
    if st.button("🔥 START MERGE", type="primary"):
        # Load Configs
        DOC_ID = st.secrets["DOC_ID"]
        SHEET_ID = st.secrets["SHEET_ID"]
        DISPLAY_NAME = st.secrets["DISPLAY_NAME"]
        accounts = json.loads(st.secrets["DUMMY_ACCOUNTS"])
        
        status_box = st.empty()
        logs = []
        
        try:
            # Get JD using first account
            admin_creds = load_creds(accounts[0])
            if not admin_creds:
                st.error("First account needs to be logged in!")
                st.stop()
                
            subject, body_template = get_jd(admin_creds, DOC_ID)
            st.info(f"Using JD: {subject}")
            
            # Loop accounts
            for i, sender in enumerate(accounts):
                creds = load_creds(sender)
                if not creds:
                    logs.append(f"❌ {sender}: Skipped (Not logged in)")
                    status_box.dataframe(logs)
                    continue
                
                # Get emails for this filter
                targets = get_sheet_col(creds, SHEET_ID, f"filter{i}")
                
                # Sending Loop
                count = 0
                for target in targets:
                    if count >= limit: break
                    
                    # Personalize
                    fname = target.split('@')[0].split('.')[0].capitalize()
                    final_body = body_template.replace("{first_name}", fname)
                    
                    try:
                        send_mail(creds, sender, target, subject, final_body, DISPLAY_NAME)
                        log_send(creds, SHEET_ID, [target, subject, sender, str(datetime.now())])
                        logs.append(f"✅ {sender} -> {target}")
                        count += 1
                        time.sleep(delay)
                    except Exception as e:
                        logs.append(f"❌ {sender} -> {target} : {e}")
                    
                    status_box.dataframe(logs)
                    
            st.success("Batch Completed!")
            
        except Exception as e:
            st.error(f"Error: {e}")
