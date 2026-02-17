import streamlit as st
import base64
import time
from datetime import datetime, timezone, timedelta
from api.gmail import fetch_all_threads, send_inbox_reply

def render():
    # Define IST offset (UTC + 5:30)
    IST = timezone(timedelta(hours=5, minutes=30))
    
    if "active_thread_id" not in st.session_state:
        st.session_state.active_thread_id = None

    if st.session_state.active_thread_id is None:
        c_head, c_btn = st.columns([4, 1])
        c_head.subheader("📥 Unified Vendor Inbox")
        if c_btn.button("🔄 Refresh Inbox", use_container_width=True):
            fetch_all_threads.clear()
            st.rerun()
        
        emails = fetch_all_threads()
        
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
                    # Apply explicit IST timezone
                    dt_obj = datetime.fromtimestamp(timestamp_ms / 1000.0, tz=IST)
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
        emails = fetch_all_threads()
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
            
            # Format individual messages to IST using Internal_Date
            try:
                msg_ts = int(msg['Internal_Date'])
                msg_dt = datetime.fromtimestamp(msg_ts / 1000.0, tz=IST)
                msg_date_str = msg_dt.strftime("%b %d, %I:%M %p")
            except:
                msg_date_str = msg["Date"]
                
            with st.container(border=True):
                if is_me: st.markdown(f"🟢 **You** (`{msg_date_str}`)")
                else: st.markdown(f"🔵 **Vendor** - {msg['From']} (`{msg_date_str}`)")
                
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
                    send_inbox_reply(
                        email=selected_thread["Account"],
                        thread_id=selected_thread["Thread_ID"],
                        rfc_message_id=selected_thread["Last_RFC_Message_ID"],
                        to_address=selected_thread["Vendor_Email"],
                        subject=selected_thread["Subject"],
                        body_text=reply_body
                    )
                st.success("Reply sent!")
                time.sleep(1)
                fetch_all_threads.clear()
                st.session_state.active_thread_id = None 
                st.rerun()
            else:
                st.error("Cannot send an empty message.")
