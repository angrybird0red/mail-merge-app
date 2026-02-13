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

# --- TAB 1: OPERATIONS CENTER (ROUND ROBIN) ---
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
            delay = st.number_input("Delay (seconds) - applied after each round", 5, 300, 20)
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
        
        # --- A. INITIALIZE DASHBOARD & PRE-LOAD DATA ---
        dashboard_data = []
        active_senders = []
        
        # Pre-load credentials and target lists to avoid API lag during sending
        with st.status("📋 Preparing Campaign...", expanded=True) as status:
            for sender in selected_accounts:
                st.write(f"Loading data for {sender}...")
                original_idx = all_accounts.index(sender)
                
                # Default dashboard row
                row = {
                    "Filter": f"Filter {original_idx}",
                    "Account": sender,
                    "Target Email": "Waiting...",
                    "Sent": 0,
                    "Errors": 0,
                    "Status": "Ready"
                }
                dashboard_data.append(row)
                
                creds = load_creds(sender)
                if creds:
                    targets = get_sheet_col(creds, SHEET_ID, f"filter{original_idx}")
                    active_senders.append({
                        "email": sender,
                        "creds": creds,
                        "targets": targets,
                        "idx": 0 # Track how many this sender has sent
                    })
                else:
                    # Mark as failed in dashboard immediately
                    for r in dashboard_data:
                        if r["Account"] == sender: r["Status"] = "❌ Auth Failed"

            try:
                # Get JD Template (once)
                admin_creds = load_creds(all_accounts[0])
                subject, body_template = get_jd(admin_creds, DOC_ID)
                status.update(label="✅ Ready to Launch!", state="complete", expanded=False)
            except Exception as e:
                st.error(f"Failed to load Template: {e}")
                st.stop()

        # Create DataFrame & UI Element
        dashboard_df = pd.DataFrame(dashboard_data).set_index("Account")
        table_placeholder = st.empty()
        table_placeholder.dataframe(dashboard_df, use_container_width=True)
        
        # --- B. ROUND ROBIN LOOP ---
        # We loop from 0 to LIMIT. In each iteration, every account sends 1 email.
        for round_num in range(limit):
            
            emails_sent_this_round = 0
            
            # 1. Send one email from EACH active sender
            for sender_obj in active_senders:
                sender_email = sender_obj["email"]
                current_idx = sender_obj["idx"]
                targets = sender_obj["targets"]
                
                # Check if this sender still has targets and hasn't hit limit
                if current_idx < len(targets):
                    target = targets[current_idx]
                    
                    # Update Dashboard: "Processing..."
                    dashboard_df.at[sender_email, "Target Email"] = target
                    dashboard_df.at[sender_email, "Status"] = "🚀 Sending..."
                    table_placeholder.dataframe(dashboard_df, use_container_width=True)
                    
                    # Personalize
                    fname = target.split('@')[0].split('.')[0].capitalize()
                    final_body = body_template.replace("{first_name}", fname)
                    
                    try:
                        if is_dry_run:
                            time.sleep(0.5) # Fake send time
                        else:
                            send_mail(sender_obj["creds"], sender_email, target, subject, final_body, DISPLAY_NAME)
                            log_send(sender_obj["creds"], SHEET_ID, [target, subject, sender_email, str(datetime.now())])
                        
                        # Success Update
                        sender_obj["idx"] += 1
                        dashboard_df.at[sender_email, "Sent"] = sender_obj["idx"]
                        dashboard_df.at[sender_email, "Status"] = f"✅ Sent ({round_num + 1})"
                        emails_sent_this_round += 1
                        
                    except Exception as e:
                        dashboard_df.at[sender_email, "Errors"] += 1
                        dashboard_df.at[sender_email, "Status"] = "❌ Error"
                    
                    # Small buffer between ACCOUNTS (to prevent API burst)
                    time.sleep(1) 
                    table_placeholder.dataframe(dashboard_df, use_container_width=True)
            
            # 2. End of Round Check
            if emails_sent_this_round == 0:
                break # Everyone ran out of targets
                
            # 3. THE BIG DELAY (Waits after the WHOLE batch is done)
            if round_num < limit - 1: # Don't wait after the very last round
                dashboard_df["Status"] = dashboard_df["Status"].apply(lambda x: f"⏳ Waiting {delay}s..." if "Sent" in str(x) else x)
                table_placeholder.dataframe(dashboard_df, use_container_width=True)
                time.sleep(delay)

        st.success("Batch Run Completed!")
