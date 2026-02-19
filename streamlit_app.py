import streamlit as st
from google_auth_oauthlib.flow import Flow

from api.auth import get_client_config, SCOPES
from ui import tab_accounts, tab_preview, tab_operations

st.set_page_config(page_title="Simple Merge", page_icon="👔", layout="wide")

# --- 1. UI SETUP & SESSION STATE ---
if 'campaign_running' not in st.session_state: st.session_state.campaign_running = False
if 'stop_clicked' not in st.session_state: st.session_state.stop_clicked = False

st.title("👔 Simple Merge")

# Catch OAuth Redirect
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
        st.stop()
    except Exception as e:
        st.error(f"Login Error: {str(e)}")

# --- 2. TABS ROUTING ---
t_run, t_preview, t_auth = st.tabs(["⚡ Operations", "👁️ Preview", "⚙️ Accounts"])

with t_auth:
    tab_accounts.render()

with t_preview:
    tab_preview.render()

with t_run:
    tab_operations.render()
