import streamlit as st
import json
from api.auth import load_creds
from api.docs import get_jd_html

def render():
    admin_email = json.loads(st.secrets["DUMMY_ACCOUNTS"])[0]
    creds = load_creds(admin_email)
    if creds:
        try:
            subj, html_body = get_jd_html(creds, st.secrets["DOC_ID"])
            st.info(f"📄 Template: {subj}")
            st.markdown("**Personalized Preview (with HTML formatting):**")
            preview_html = html_body.replace("{first_name}", "John").replace("{company}", "TechCorp").replace("{job_title}", "Analyst")
            st.html(preview_html)
        except Exception as e: st.error(f"Could not load preview: {e}")
    else: st.warning("Connect your first account to preview.")
