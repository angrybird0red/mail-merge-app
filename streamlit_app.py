import streamlit as st
import json
import time
import random
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

st.set_page_config(page_title="Mail Merge Elite V5", page_icon="👔", layout="wide")

# --- 1. CORE LOGIC (FORMATTING FIX) ---
# We use Drive API for HTML export to keep 1:1 formatting
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
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
        token_info = json.loads(st.secrets[secret_key])
        creds = Credentials.from_authorized_user_info(token_info)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    return None

def get_jd_html(creds, doc_id):
    """Exports Google Doc as HTML to preserve 1:1 formatting."""
    drive_service = build('drive', 'v3', credentials=creds)
    html_content = drive_service.files().export(fileId=doc_id, mimeType='text/html').execute()
    
    docs_service = build('docs', 'v1', credentials=creds)
    doc = docs_service.documents().get(documentId=doc_id).execute()
    return doc.get('title'), html_content.decode('utf-8')

def get_full_sheet_data(creds, sheet_id, sheet_name):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        return res.get('values', []) if res.get('values') else []
    except: return []

def send_mail_html(creds, sender, to, subject, html_body, display_name):
    """Sends a professional HTML-formatted email."""
    service = build('gmail', 'v1', credentials=creds)
    msg = EmailMessage()
    
    # Set the content subtype to 'html' to render bold/line breaks
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(html_body)
    
    msg['To'] = to
    msg['From'] = formataddr((display_name, sender))
    msg['Subject'] = subject
    
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={'raw': raw}).execute()

# --- 2. SESSION STATE ---
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False

# --- 3. UI TABS ---
st.title("👔 Mail Merge Elite V5")
tab_run, tab_preview, tab_auth = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts"])

# --- TAB: ACCOUNTS (FIXED FOR RE-AUTH) ---
with tab_auth:
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    
    # ... (Keep your query_params/code handling block here) ...

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        # CHANGED: We removed "if not creds:" so the button is ALWAYS there
        if col2.button("Login / Refresh", key=f"login_{email}"):
            redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
            flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
            # The 'state' ensures Google tells the app EXACTLY which account is signing in
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
            # Replace tags in the HTML string
            preview_html = html_body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
            # Render the HTML in Streamlit for preview
            st.html(preview_html)
        except Exception as e: st.error(f"Could not load preview: {e}")
    else: st.warning("Connect your first account to preview.")

# --- TAB: OPERATIONS ---
with tab_run:
    all_acc = json.loads(st.secrets["DUMMY_ACCOUNTS"])
    with st.sidebar:
        st.header("⚙️ Configuration")
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
                    
                    # The formatting fix:
                    final_body = body_tmpl.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                    
                    dashboard_df.at[s["email"], "Target"] = target
                    dashboard_df.at[s["email"], "Status"] = "📨 Sending..."
                    table_ui.dataframe(dashboard_df, use_container_width=True)

                    try:
                        if not is_dry: send_mail_html(s["creds"], s["email"], target, subj, final_body, st.secrets["DISPLAY_NAME"])
                        s["idx"] += 1
                        sent_total += 1
                        dashboard_df.at[s["email"], "Sent"] = s["idx"]
                        dashboard_df.at[s["email"], "Status"] = "✅ Sent"
                    except Exception as e:
                        dashboard_df.at[s["email"], "Status"] = f"❌ Error"
                    
                    time.sleep(1)
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    prog_bar.progress(sent_total/total_goal, text=f"Sent {sent_total} / {total_goal}")

            if not round_active: break
            
            if r_num < limit - 1:
                human_delay = delay + random.randint(-2, 2)
                for sec in range(human_delay, 0, -1):
                    if st.session_state.stop_clicked: break
                    for s in active_data: dashboard_df.at[s["email"], "Status"] = f"⏳ {sec}s"
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    time.sleep(1)

        st.balloons()
