import streamlit as st
import json
import time
import random
import pandas as pd
import threading
from streamlit.runtime.scriptrunner import add_script_run_ctx

from api.auth import load_creds
from api.docs import get_jd_html
from api.sheets import get_send_log, get_full_sheet_data, append_to_send_log
from api.gmail import send_mail_html

# --- CALLBACK FUNCTION ---
def stop_campaign():
    st.session_state.stop_clicked = True

def render():
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
    
    # Updated to use on_click callback instead of if-statement
    col_stop.button("🛑 STOP CAMPAIGN", type="secondary", disabled=not st.session_state.campaign_running, on_click=stop_campaign)

    if start and not st.session_state.campaign_running:
        st.session_state.stop_clicked = False
        st.session_state.campaign_running = True
        
        with st.status("🔍 Initializing System...") as status:
            admin_creds = load_creds(all_acc[0])
            subj, body_tmpl = get_jd_html(admin_creds, st.secrets["DOC_ID"])
            
            st.session_state.sent_history = get_send_log(admin_creds, st.secrets["SHEET_ID"])
            
            active_data = []
            for s in sel_acc:
                c = load_creds(s)
                if c:
                    rows = get_full_sheet_data(c, st.secrets["SHEET_ID"], f"filter{all_acc.index(s)}")
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
                                    send_mail_html(s["creds"], s["email"], target, subj, final_body, display_name)
                                    append_to_send_log(s["creds"], st.secrets["SHEET_ID"], target, subj, s["email"])
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
        st.rerun() # Forces immediate UI update so buttons correctly toggle

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
