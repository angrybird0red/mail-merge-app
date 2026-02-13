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

st.set_page_config(page_title="Mail Merge Elite V4", page_icon="👔", layout="wide")

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
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        return res.get('values', []) if res.get('values') else []
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

# --- 2. SESSION STATE ---
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False

# --- 3. UI TABS ---
st.title("👔 Mail Merge Elite V4")
tab_run, tab_preview, tab_auth = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts"])

# --- TAB: ACCOUNTS ---
with tab_auth:
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))
    
    # 1. HANDLE REDIRECT LOGIN (Success Screen)
    if "code" in st.query_params:
        code, email_trying = st.query_params["code"], st.query_params.get("state")
        try:
            redirect_uri = "https://mail-merge-app-angrybird0red.streamlit.app"
            flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
            flow.fetch_token(code=code)
            st.success(f"✅ Success for {email_trying}")
            st.code(flow.credentials.to_json(), language="json")
            st.info(f"Copy/Paste above into Secrets as `TOKEN_{email_trying.replace('@','_').replace('.','_').upper()}`")
        except Exception as e: st.error(str(e))

    # 2. SHOW ALL ACCOUNTS & LOGIN BUTTONS
    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        if not creds: # ALWAYS show login button if not connected
            if col2.button("Login", key=f"login_{email}"):
                redirect_uri = "https://mail-merge-app-angrybird0red.streamlit.app"
                flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
                url, _ = flow.authorization_url(prompt='consent', state=email)
                st.link_button("👉 Start Auth", url)

# --- TAB: PREVIEW ---
with tab_preview:
    admin_email = json.loads(st.secrets["DUMMY_ACCOUNTS"])[0]
    creds = load_creds(admin_email)
    if creds:
        subj, body = get_jd(creds, st.secrets["DOC_ID"])
        st.info(f"📄 Template: {subj}")
        st.markdown("**Example Personalization:**")
        preview = body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
        st.text_area("", preview, height=300)
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
            subj, body_tmpl = get_jd(admin_creds, st.secrets["DOC_ID"])
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

        # --- ROUND ROBIN EXECUTION ---
        for r_num in range(limit):
            if st.session_state.stop_clicked:
                st.error("🛑 Stop button pressed. Terminating campaign...")
                break
            
            round_active = False
            for s in active_data:
                if s["idx"] < len(s["rows"]):
                    round_active = True
                    row = s["rows"][s[ "idx"]]
                    target = row[0]
                    
                    # Mapping Variables
                    comp = row[1] if len(row) > 1 else "Your Company"
                    role = row[2] if len(row) > 2 else "the open position"
                    fname = target.split('@')[0].split('.')[0].capitalize()
                    
                    final_body = body_tmpl.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                    
                    dashboard_df.at[s["email"], "Target"] = target
                    dashboard_df.at[s["email"], "Status"] = "📨 Sending..."
                    table_ui.dataframe(dashboard_df, use_container_width=True)

                    try:
                        if not is_dry: send_mail(s["creds"], s["email"], target, subj, final_body, st.secrets["DISPLAY_NAME"])
                        s["idx"] += 1
                        sent_total += 1
                        dashboard_df.at[s["email"], "Sent"] = s["idx"]
                        dashboard_df.at[s["email"], "Status"] = "✅ Sent"
                    except: dashboard_df.at[s["email"], "Status"] = "❌ Error"
                    
                    time.sleep(1) # Tiny delay between accounts
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    prog_bar.progress(sent_total/total_goal, text=f"Sent {sent_total} / {total_goal}")

            if not round_active: break
            
            # THE COUNTDOWN
            if r_num < limit - 1:
                human_delay = delay + random.randint(-2, 2) # HUMAN-LIKE RANDOMIZATION
                for sec in range(human_delay, 0, -1):
                    if st.session_state.stop_clicked: break
                    for s in active_data: dashboard_df.at[s["email"], "Status"] = f"⏳ {sec}s"
                    table_ui.dataframe(dashboard_df, use_container_width=True)
                    time.sleep(1)

        st.balloons()
