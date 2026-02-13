import streamlit as st
import json
import time
import random
import pandas as pd
from datetime import datetime
import base64

# --- NEW IMPORTS FOR ROBUST HTML EMAILS ---
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr

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
    # Standardize key format
    safe_email = email.replace("@", "_").replace(".", "_").upper()
    secret_key = f"TOKEN_{safe_email}"
    
    if secret_key in st.secrets:
        try:
            # First load: Get the data from Secrets
            token_data = st.secrets[secret_key]
            
            # 1. Handle JSON parsing
            try:
                token_info = json.loads(token_data)
            except:
                # If it's not JSON, maybe it's already a dict?
                token_info = token_data
            
            # 2. THE FIX: Double-Check if it's still a string (Double-Encoded)
            if isinstance(token_info, str):
                token_info = json.loads(token_info)

            # 3. Auto-Fill Missing Client ID/Secret (The Safety Net)
            if isinstance(token_info, dict) and "client_id" not in token_info:
                main_config = json.loads(st.secrets["gcp_service_account"])
                app_info = main_config.get("web", main_config.get("installed", {}))
                token_info["client_id"] = app_info.get("client_id")
                token_info["client_secret"] = app_info.get("client_secret")
            
            # 4. Create Credentials
            creds = Credentials.from_authorized_user_info(token_info)
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
            return creds
            
        except Exception as e:
            st.error(f"⚠️ Token Error for {email}: {e}")
            return None
    return None

def get_jd_html(creds, doc_id):
    """Exports Google Doc as HTML and aggressively fixes spacing issues."""
    drive_service = build('drive', 'v3', credentials=creds)
    html_content = drive_service.files().export(fileId=doc_id, mimeType='text/html').execute()
    decoded_html = html_content.decode('utf-8')
    
    # --- CSS HACK V2: THE NUCLEAR OPTION ---
    style_fix = """
    <style>
        body { 
            font-family: Arial, Helvetica, sans-serif !important; 
            font-size: 11pt !important; 
            line-height: 1.4 !important; 
            color: #000000 !important;
            background-color: transparent !important;
        }
        p { margin-top: 0 !important; margin-bottom: 12px !important; padding: 0 !important; }
        ul, ol { margin-top: 0 !important; margin-bottom: 12px !important; padding-left: 30px !important; }
        li { margin-bottom: 2px !important; margin-top: 0 !important; }
        li p { margin: 0 !important; padding: 0 !important; display: inline !important; }
        h1, h2, h3, h4 { margin-top: 18px !important; margin-bottom: 8px !important; font-weight: bold !important; }
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
    """Sends a professional HTML-formatted email using MIMEMultipart."""
    service = build('gmail', 'v1', credentials=creds)
    
    # Use MIMEMultipart - this is the standard for HTML emails
    message = MIMEMultipart("alternative")
    message['To'] = to
    message['From'] = formataddr((display_name, sender))
    message['Subject'] = subject

    # Attach the HTML body properly
    msg_html = MIMEText(html_body, "html")
    message.attach(msg_html)
    
    # Encode and send
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId="me", body={'raw': raw}).execute()

# --- 2. UI SETUP ---
st.title("👔 Mail Merge Elite V5")

# [[[ CHECK FOR LOGIN SUCCESS AT THE VERY TOP ]]]
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
        st.stop() # Stop loading the rest of the app so you focus on copying
    except Exception as e:
        st.error(f"Login Error: {str(e)}")

# --- 3. TABS ---
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False
tab_run, tab_preview, tab_auth = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts"])

# --- TAB: ACCOUNTS ---
with tab_auth:
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        # Always show login button to allow re-auth/refresh
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
            # Replace tags in the HTML string
            preview_html = html_body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
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
                    
                    final_body = body_tmpl.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                    
                    dashboard_df.at[s["email"], "Target"] = target
                    dashboard_df.at[s["email"], "Status"] = "📨 Sending..."
                    table_ui.dataframe(dashboard_df, use_container_width=True)

                    try:
                        if not is_dry: 
                            send_mail_html(s["creds"], s["email"], target, subj, final_body, st.secrets["DISPLAY_NAME"])
                        
                        s["idx"] += 1
                        sent_total += 1
                        dashboard_df.at[s["email"], "Sent"] = s["idx"]
                        dashboard_df.at[s["email"], "Status"] = "✅ Sent"
                    except Exception as e:
                        # --- ERROR HANDLING FIX ---
                        # Indented correctly now!
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
