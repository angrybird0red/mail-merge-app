import streamlit as st
import json
import time
import random
import pandas as pd
import threading
import base64
from datetime import datetime
from streamlit.runtime.scriptrunner import add_script_run_ctx
from google_auth_oauthlib.flow import Flow

# IMPORT LOCAL MODULES
import utils

st.set_page_config(page_title="Simple Merge", page_icon="👔", layout="wide")

# --- 1. UI SETUP & SESSION STATE ---
if 'campaign_running' not in st.session_state: st.session_state.campaign_running = False
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False

st.title("👔 Simple Merge")

if "code" in st.query_params:
    code = st.query_params["code"]
    email_trying = st.query_params.get("state", "Unknown Account")
    try:
        redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
        flow = Flow.from_client_config(utils.get_client_config(), utils.SCOPES, redirect_uri=redirect_uri)
        flow.fetch_token(code=code)
        
        st.success(f"✅ LOGIN SUCCESS FOR: {email_trying}")
        st.warning("⬇️ COPY THIS TOKEN BELOW AND PASTE INTO SECRETS ⬇️")
        st.code(flow.credentials.to_json(), language="json")
        st.stop()
    except Exception as e:
        st.error(f"Login Error: {str(e)}")

# --- 2. TABS ---
tab_run, tab_preview, tab_auth, tab_inbox = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts", "📥 Inbox"])

# --- TAB: ACCOUNTS ---
with tab_auth:
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = utils.load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        if col2.button("Login / Refresh", key=f"login_{email}"):
            redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
            flow = Flow.from_client_config(utils.get_client_config(), utils.SCOPES, redirect_uri=redirect_uri)
            url, _ = flow.authorization_url(prompt='consent', state=email)
            st.link_button("👉 Start Auth", url)

# --- TAB: PREVIEW ---
with tab_preview:
    admin_email = json.loads(st.secrets["DUMMY_ACCOUNTS"])[0]
    creds = utils.load_creds(admin_email)
    if creds:
        try:
            subj, html_body = utils.get_jd_html(creds, st.secrets["DOC_ID"])
            st.info(f"📄 Template: {subj}")
            st.markdown("**Personalized Preview (with HTML formatting):**")
            preview_html = html_body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
            st.html(preview_html)
        except Exception as e: st.error(f"Could not load preview: {e}")
    else: st.warning("Connect your first account to preview.")

# --- TAB: OPERATIONS (BACKGROUND THREADING) ---
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
    start = col_btn.button("🔥 LAUNCH", type="primary", use_container_width=True, disabled=st.session_state.campaign_running)
    if col_stop.button("🛑 STOP CAMPAIGN", type="secondary", disabled=not st.session_state.campaign_running): 
        st.session_state.stop_clicked = True

    if start and not st.session_state.campaign_running:
        st.session_state.stop_clicked = False
        st.session_state.campaign_running = True
        
        with st.status("🔍 Initializing System...") as status:
            admin_creds = utils.load_creds(all_acc[0])
            subj, body_tmpl = utils.get_jd_html(admin_creds, st.secrets["DOC_ID"])
            
            st.session_state.sent_history = utils.get_send_log(admin_creds, st.secrets["SHEET_ID"])
            
            active_data = []
            for s in sel_acc:
                c = utils.load_creds(s)
                if c:
                    rows = utils.get_full_sheet_data(c, st.secrets["SHEET_ID"], f"filter{all_acc.index(s)}")
                    active_data.append({"email": s, "creds": c, "rows": rows, "idx": 0, "sent_count": 0})
            status.update(label="System Ready! Handing off to background.", state="complete")

        st.session_state.active_data = active_data
        st.session_state.dashboard_df = pd.DataFrame([{"Account": s["email"], "Target": "-", "Sent": 0, "Status": "Ready"} for s in active_data]).set_index("Account")
        st.session_state.total_goal = sum([min(len(s["rows"]), limit) for s in active_data])
        st.session_state.sent_total = 0

        def background_campaign(limit, delay, is_dry, display_name, subj, body_tmpl):
            try:
                for s in st.session_state.active_data:
                    s["sent_count"] = 0

                while True:
                    if st.session_state.stop_clicked: break
                    
                    round_active = False
                    for s in st.session_state.active_data:
                        # Fast-forward past duplicates
                        while s["idx"] < len(s["rows"]) and s["sent_count"] < limit:
                            row = s["rows"][s["idx"]]
                            target = row[0].strip()
                            
                            if (target, subj) in st.session_state.sent_history:
                                st.session_state.dashboard_df.at[s["email"], "Status"] = "⏭️ Skipped"
                                s["idx"] += 1
                                time.sleep(0.1)
                                continue 
                            
                            round_active = True
                            comp = row[1] if len(row) > 1 else "Your Company"
                            role = row[2] if len(row) > 2 else "the open position"
                            fname = target.split('@')[0].split('.')[0].capitalize()
                            
                            st.session_state.dashboard_df.at[s["email"], "Target"] = target
                            st.session_state.dashboard_df.at[s["email"], "Status"] = "📨 Sending..."

                            try:
                                if not is_dry: 
                                    final_body = body_tmpl.replace("{first_name}", fname).replace("{company}", comp).replace("{job_title}", role)
                                    utils.send_mail_html(s["creds"], s["email"], target, subj, final_body, display_name)
                                    utils.append_to_send_log(s["creds"], st.secrets["SHEET_ID"], target, subj, s["email"])
                                    st.session_state.sent_history.add((target, subj))
                                
                                s["idx"] += 1
                                s["sent_count"] += 1
                                st.session_state.sent_total += 1
                                st.session_state.dashboard_df.at[s["email"], "Sent"] = s["sent_count"]
                                st.session_state.dashboard_df.at[s["email"], "Status"] = "✅ Sent"
                            except Exception as e:
                                error_msg = str(e).split(']')[0] 
                                st.session_state.dashboard_df.at[s["email"], "Status"] = f"❌ {error_msg}"
                                s["idx"] += 1 
                            
                            time.sleep(1)
                            break 

                    if not round_active: break
                    
                    human_delay = delay + random.randint(-2, 2)
                    for sec in range(human_delay, 0, -1):
                        if st.session_state.stop_clicked: break
                        for s in st.session_state.active_data: 
                            status_val = str(st.session_state.dashboard_df.at[s["email"], "Status"])
                            if "Auth" not in status_val and "Error" not in status_val and "Skipped" not in status_val:
                                st.session_state.dashboard_df.at[s["email"], "Status"] = f"⏳ {sec}s"
                        time.sleep(1)
            finally:
                st.session_state.campaign_running = False

        thread = threading.Thread(target=background_campaign, args=(limit, delay, is_dry, display_name, subj, body_tmpl))
        add_script_run_ctx(thread)
        thread.start()

    if st.session_state.campaign_running:
        @st.fragment(run_every="2s")
        def render_campaign_progress():
            if st.session_state.campaign_running:
                st.info("🚀 Campaign is running in the background. You can safely switch to the Inbox tab!")
                prog = st.session_state.sent_total / max(1, st.session_state.total_goal)
                st.progress(prog, text=f"Sent {st.session_state.sent_total} / {st.session_state.total_goal}")
                st.dataframe(st.session_state.dashboard_df, use_container_width=True)
            else:
                st.rerun()
        render_campaign_progress()
        
    elif getattr(st.session_state, 'sent_total', 0) > 0 and not st.session_state.stop_clicked:
        st.success("✅ Campaign Completed!")
        st.progress(1.0, text=f"Sent {st.session_state.sent_total} / {st.session_state.total_goal}")
        st.dataframe(st.session_state.dashboard_df, use_container_width=True)
        
    elif st.session_state.stop_clicked and hasattr(st.session_state, 'dashboard_df'):
        st.warning("🛑 Campaign Stopped.")
        st.dataframe(st.session_state.dashboard_df, use_container_width=True)

# --- TAB: INBOX ---
with tab_inbox:
    if "active_thread_id" not in st.session_state:
        st.session_state.active_thread_id = None

    if st.session_state.active_thread_id is None:
        c_head, c_btn = st.columns([4, 1])
        c_head.subheader("📥 Unified Vendor Inbox")
        if c_btn.button("🔄 Refresh Inbox", use_container_width=True):
            utils.fetch_all_threads.clear()
            st.rerun()
        
        emails = utils.fetch_all_threads()
        
        if not emails:
            st.info("No active vendor conversations found.")
        else:
            st.divider()
            h_col1, h_col2, h_col3, h_col4, h_col5 = st.columns([2, 4, 2, 2, 1])
            h_col1.markdown("**Vendor**")
            h_col2.markdown("**Subject & Preview**")
            h_col3.markdown("**Account**")
            h_col4.markdown("**Date/Time**")
            h_col5.markdown("**Action**")
            
            for em in emails:
                st.markdown("<hr style='margin: 0px; padding: 0px; border-top: 1px solid #ddd;'>", unsafe_allow_html=True)
                c1, c2, c3, c4, c5 = st.columns([2, 4, 2, 2, 1])
                
                sender = em['Vendor_Email'].split('<')[0].strip()[:20]
                subj = em['Subject'][:40]
                acc = em['Account'].split('@')[0]
                snippet = em['Messages'][-1].get('Snippet', '')[:50]
                
                try:
                    timestamp_ms = int(em['Last_Message_Time'])
                    dt_obj = datetime.fromtimestamp(timestamp_ms / 1000.0)
                    formatted_date = dt_obj.strftime("%b %d, %I:%M %p")
                except:
                    formatted_date = "Unknown"
                
                c1.markdown(f"<div style='padding-top: 8px;'>👤 {sender}</div>", unsafe_allow_html=True)
                c2.markdown(f"<div style='padding-top: 8px;'><b>{subj}</b> - <span style='color: gray;'>{snippet}...</span></div>", unsafe_allow_html=True)
                c3.markdown(f"<div style='padding-top: 8px;'>📥 {acc}</div>", unsafe_allow_html=True)
                c4.markdown(f"<div style='padding-top: 8px;'>🕒 {formatted_date}</div>", unsafe_allow_html=True)
                
                if c5.button("Open", key=f"open_{em['Thread_ID']}", use_container_width=True):
                    st.session_state.active_thread_id = em['Thread_ID']
                    st.rerun()
            st.markdown("<hr style='margin: 0px; padding: 0px; border-top: 1px solid #ddd;'>", unsafe_allow_html=True)

    else:
        emails = utils.fetch_all_threads()
        selected_thread = next((t for t in emails if t['Thread_ID'] == st.session_state.active_thread_id), None)
        
        if not selected_thread:
            st.session_state.active_thread_id = None
            st.rerun()

        if st.button("⬅️ Back to Inbox"):
            st.session_state.active_thread_id = None
            st.rerun()
            
        st.markdown(f"### {selected_thread['Subject']}")
        st.caption(f"**Vendor:** `{selected_thread['Vendor_Email']}` | **Via:** `{selected_thread['Account']}`")
        st.divider()
        
        for msg in selected_thread["Messages"]:
            is_me = selected_thread["Account"] in msg["From"]
            with st.container(border=True):
                if is_me: st.markdown(f"🟢 **You** (`{msg['Date']}`)")
                else: st.markdown(f"🔵 **Vendor** - {msg['From']} (`{msg['Date']}`)")
                
                st.html(msg["Body"])
                
                if msg["Attachments"]:
                    st.markdown("**Attachments:**")
                    for att in msg["Attachments"]:
                        col_a, col_b = st.columns([3, 1])
                        with col_a: st.write(f"📎 `{att['filename']}`")
                        with col_b: st.download_button(label="Download", data=att['data'], file_name=att['filename'], key=f"dl_{msg['Message_ID']}_{att['filename']}")
                        
                        with st.expander(f"👁️ Preview {att['filename']}"):
                            fname_lower = att['filename'].lower()
                            try:
                                if fname_lower.endswith('.pdf'):
                                    base64_pdf = base64.b64encode(att['data']).decode('utf-8')
                                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="500" type="application/pdf"></iframe>'
                                    st.markdown(pdf_display, unsafe_allow_html=True)
                                elif fname_lower.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                                    st.image(att['data'])
                                elif fname_lower.endswith(('.txt', '.csv', '.xml', '.json', '.md', '.html')):
                                    st.code(att['data'].decode('utf-8'))
                                elif fname_lower.endswith(('.docx', '.doc')):
                                    st.info("Microsoft Word formats (.docx, .doc) cannot be previewed directly. Please download.")
                                else:
                                    st.info("Preview not available.")
                            except Exception as e:
                                st.warning("Could not generate a preview.")

        st.divider()
        st.markdown("#### Quick Reply")
        reply_body = st.text_area("Message:", key=f"reply_{selected_thread['Thread_ID']}")
        
        if st.button("Send Reply", type="primary", key=f"btn_{selected_thread['Thread_ID']}"):
            if reply_body.strip():
                with st.spinner("Sending..."):
                    utils.send_inbox_reply(
                        email=selected_thread["Account"],
                        thread_id=selected_thread["Thread_ID"],
                        rfc_message_id=selected_thread["Last_RFC_Message_ID"],
                        to_address=selected_thread["Vendor_Email"],
                        subject=selected_thread["Subject"],
                        body_text=reply_body
                    )
                st.success("Reply sent!")
                time.sleep(1)
                utils.fetch_all_threads.clear()
                st.session_state.active_thread_id = None 
                st.rerun()
            else:
                st.error("Cannot send an empty message.")
