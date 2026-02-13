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

st.set_page_config(page_title="Mail Merge Pro", page_icon="🚀", layout="wide")

# --- 1. SETUP & AUTH ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

def get_client_config():
    if "gcp_service_account" in st.secrets:
        return json.loads(st.secrets["gcp_service_account"])
    else:
        st.error("🚨 Missing 'gcp_service_account' in Secrets!")
        st.stop()

def load_creds(email):
    # Standardize key format: TOKEN_USER_GMAIL_COM
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
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:A").execute()
        return [row[0] for row in res.get('values', []) if row]
    except Exception:
        return []

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

# --- 3. THE UI ---
st.title("🚀 Mail Merge Pro")

# TABS
tab_run, tab_auth = st.tabs(["⚡ Operations Center", "⚙️ Account Manager"])

# --- TAB 2: ACCOUNT MANAGER ---
with tab_auth:
    st.write("### Connected Accounts")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    
    if "code" in st.query_params:
        code = st.query_params["code"]
        email_trying = st.query_params.get("state")
        try:
            redirect_uri = "https://mail-merge-app-angrybird0red.streamlit.app"
            flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.success(f"✅ Authenticated: {email_trying}")
            st.code(creds.to_json(), language="json")
            st.info(f"Save as: `TOKEN_{email_trying.replace('@','_').replace('.','_').upper()}` in Secrets")
        except Exception as e:
            st.error(str(e))

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        if not creds:
            if col2.button("Login", key=email):
                redirect_uri = "https://mail-merge-app-angrybird0red.streamlit.app"
                flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
                url, _ = flow.authorization_url(prompt='consent', state=email)
                st.link_button("👉 Sign In", url)

# --- TAB 1: OPERATIONS CENTER (FIXED DASHBOARD) ---
with tab_run:
    # Load Configs
    DOC_ID = st.secrets.get("DOC_ID")
    SHEET_ID = st.secrets.get("SHEET_ID")
    DISPLAY_NAME = st.secrets.get("DISPLAY_NAME", "Recruitment Team")
    all_accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))

    # 1. CONTROL PANEL
    with st.container(border=True):
        st.subheader("🎛️ Control Panel")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            selected_accounts = st.multiselect("Select Senders", all_accounts, default=all_accounts)
        with c2:
            limit = st.number_input("Max Emails per Account", 1, 500, 20)
            delay = st.number_input("Delay (seconds)", 5, 120, 20)
        with c3:
            st.write("**Safety Mode**")
            is_dry_run = st.toggle("🧪 Dry Run (Test Mode)", value=True)
            if is_dry_run:
                st.info("Simulation only.")
            else:
                st.warning("⚠️ LIVE SENDING")

    # 2. EXECUTION LOGIC
    st.divider()
    if st.button("🔥 START CAMPAIGN", type="primary", disabled=not selected_accounts):
        
        # --- A. INITIALIZE DASHBOARD (The Fix) ---
        # Create a dictionary for the dashboard state
        dashboard_data = []
        for sender in selected_accounts:
            # Find which filter number this account maps to
            original_idx = all_accounts.index(sender)
            dashboard_data.append({
                "Filter": f"Filter {original_idx}",
                "Account": sender,
                "Target Email": "Waiting...",
                "Sent": 0,
                "Errors": 0,
                "Status": "Ready"
            })
        
        # Create the DataFrame and the empty container
        dashboard_df = pd.DataFrame(dashboard_data).set_index("Account")
        table_placeholder = st.empty()
        table_placeholder.dataframe(dashboard_df, use_container_width=True)
        
        try:
            # Get JD
            admin_creds = load_creds(all_accounts[0])
            subject, body_template = get_jd(admin_creds, DOC_ID)
            
            # --- B. PROCESSING LOOP ---
            for sender in selected_accounts:
                original_index = all_accounts.index(sender)
                
                # Update Status: Starting
                dashboard_df.at[sender, "Status"] = "🚀 Starting..."
                table_placeholder.dataframe(dashboard_df, use_container_width=True)
                
                creds = load_creds(sender)
                if not creds:
                    dashboard_df.at[sender, "Status"] = "❌ Auth Failed"
                    table_placeholder.dataframe(dashboard_df, use_container_width=True)
                    continue
                
                # Fetch targets
                targets = get_sheet_col(creds, SHEET_ID, f"filter{original_index}")
                
                count = 0
                for target in targets:
                    if count >= limit: 
                        dashboard_df.at[sender, "Status
