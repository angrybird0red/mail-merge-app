import streamlit as st
import json
from google_auth_oauthlib.flow import Flow
from api.auth import load_creds, get_client_config, SCOPES

def render():
    st.subheader("Account Authorization")
    accounts = json.loads(st.secrets.get("DUMMY_ACCOUNTS", "[]"))

    for email in accounts:
        col1, col2 = st.columns([3, 1])
        creds = load_creds(email)
        status = "✅ Ready" if creds else "❌ Disconnected"
        col1.write(f"**{email}** : {status}")
        
        if col2.button("Login / Refresh", key=f"login_{email}"):
            redirect_uri = "https://mail-merge-app-xuxkqmkhigxrnyoeftbfif.streamlit.app"
            flow = Flow.from_client_config(get_client_config(), SCOPES, redirect_uri=redirect_uri)
            url, _ = flow.authorization_url(prompt='consent', state=email)
            st.link_button("👉 Start Auth", url)
