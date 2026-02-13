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

st.set_page_config(page_title="Mail Merge Elite", page_icon="👔", layout="wide")

# --- 1. CORE LOGIC ---
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/documents.readonly', 'https://www.googleapis.com/auth/spreadsheets']

def get_client_config():
    return json.loads(st.secrets["gcp_service_account"])

def load_creds(email):
    safe_email = email.replace("@", "_").replace(".", "_").upper()
    secret_key = f"TOKEN_{safe_email}"
    if secret_key in st.secrets:
        token_info = json.loads(st.secrets[secret_key])
        creds = Credentials.from_authorized_user_info(token_info)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    return None

def get_jd(creds, doc_id):
    docs = build('docs', 'v1', credentials=creds)
    doc = docs.documents().get(documentId=doc_id).execute()
    title = doc.get('title')
    text = "".join([e['textRun'].get('content', '') for elem in doc.get('body', {}).get('content', []) if 'paragraph' in elem for e in elem['paragraph']['elements'] if 'textRun' in e])
    return title, text

def get_full_sheet_data(creds, sheet_id, sheet_name):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        # Fetch A:C (Email, Company, Job Title)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        values = res.get('values', [])
        return values if values else []
    except: return []

def send_mail(creds, sender, to, subject, body, display_name):
    service = build('gmail', 'v1', credentials=creds)
    msg = EmailMessage()
    msg.set_content(body)
    msg['To'] = to
    msg['From'] = formataddr((display_name, sender))
    msg['Subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={'raw': raw}).execute()

# --- 2. UI TABS ---
st.title("👔 Mail Merge Elite")
tab_run, tab_preview, tab_auth = st.tabs(["⚡ Operations", "👁️ Template Preview", "⚙️ Accounts"])

# --- TAB: ACCOUNTS ---
with tab_auth:
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    # Auth logic same as previous success
    for email in accounts:
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        st.write(f"**{email}** : {status}")

# --- TAB: TEMPLATE PREVIEW ---
with tab_preview:
    st.subheader("Template Inspector")
    admin_email = json.loads(st.secrets["DUMMY_ACCOUNTS"])[0]
    creds = load_creds(admin_email)
    if creds:
        subj, body = get_jd(creds, st.secrets["DOC_ID"])
        st.info(f"📄 **Current Doc:** {subj}")
        
        # Simulation Data
        test_vars = {"first_name": "John", "company": "Google", "job_title": "Software Engineer"}
        preview_body = body
        for k, v in test_vars.items():
            preview_body = preview_body.replace(f"{{{k}}}", v)
            
        st.markdown("**Subject:**")
        st.code(subj)
        st.markdown("**Body Preview (Personalized):**")
        st.text_area("", preview_body, height=300)
    else:
        st.error("Connect your first account to see previews.")

# --- TAB: OPERATIONS ---
with tab_run:
    all_accounts = json.loads(st.secrets["DUMMY_ACCOUNTS"])
    
    with st.sidebar:
        st.header("Campaign Settings")
        selected_accounts = st.multiselect("Active Senders", all_accounts, default=all_accounts)
        limit = st.number_input("Emails per Account", 1, 1000, 20)
        delay = st.number_input("Round Delay (sec)", 5, 600, 20)
        is_dry_run = st.toggle("🧪 Dry Run", value=True)

    if st.button("🚀 LAUNCH CAMPAIGN", type="primary", use_container_width=True):
        active_data = []
        with st.status("🔍 Pre-Flight Health Check...") as status:
            admin_creds = load_creds(all_accounts[0])
            subj, body_template = get_jd(admin_creds, st.secrets["DOC_ID"])
            
            for sender in selected_accounts:
                creds = load_creds(sender)
                if creds:
                    original_idx = all_accounts.index(sender)
                    data = get_full_sheet_data(creds, st.secrets["SHEET_ID"], f"filter{original_idx}")
                    active_data.append({"email": sender, "creds": creds, "rows": data, "idx": 0})
            status.update(label="System Ready!", state="complete")

        dashboard_df = pd.DataFrame([{"Account": s["email"], "Target": "-", "Sent": 0, "Status": "Ready"} for s in active_data]).set_index("Account")
        progress_bar = st.progress(0, text="Campaign Progress")
        table_placeholder = st.empty()
        
        total_targets = sum([min(len(s["rows"]), limit) for s in active_data])
        sent_total = 0

        for round_num in range(limit):
            round_active = False
            for s_obj in active_data:
                if s_obj["idx"] < len(s_obj["rows"]):
                    round_active = True
                    row = s_obj["rows"][s_obj["idx"]]
                    target_email = row[0]
                    # Dynamic Variable Logic
                    comp = row[1] if len(row) > 1 else "Your Company"
                    role = row[2] if len(row) > 2 else "the open position"
                    fname = target_email.split('@')[0].split('.')[0].capitalize()
                    
                    final_body = body_template.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                    
                    dashboard_df.at[s_obj["email"], "Target"] = target_email
                    dashboard_df.at[s_obj["email"], "Status"] = "📨 Sending..."
                    table_placeholder.dataframe(dashboard_df)

                    try:
                        if not is_dry_run:
                            send_mail(s_obj["creds"], s_obj["email"], target_email, subj, final_body, st.secrets["DISPLAY_NAME"])
                        
                        s_obj["idx"] += 1
                        sent_total += 1
                        dashboard_df.at[s_obj["email"], "Sent"] = s_obj["idx"]
                        dashboard_df.at[s_obj["email"], "Status"] = "✅ Sent"
                        progress_bar.progress(sent_total/total_targets, text=f"Sent {sent_total} of {total_targets}")
                    except:
                        dashboard_df.at[s_obj["email"], "Status"] = "❌ Error"
                    
                    time.sleep(1)
                    table_placeholder.dataframe(dashboard_df)

            if not round_active: break
            
            if round_num < limit - 1:
                for sec in range(delay, 0, -1):
                    for s in active_data:
                        if "Auth" not in dashboard_df.at[s["email"], "Status"]:
                            dashboard_df.at[s["email"], "Status"] = f"⏳ {sec}s"
                    table_placeholder.dataframe(dashboard_df)
                    time.sleep(1)

        st.balloons()
