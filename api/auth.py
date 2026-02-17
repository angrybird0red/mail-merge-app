import streamlit as st
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# --- CONSTANTS ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
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
        try:
            token_data = st.secrets[secret_key]
            try:
                token_info = json.loads(token_data)
            except:
                token_info = token_data
           
            if isinstance(token_info, str):
                token_info = json.loads(token_info)

            if isinstance(token_info, dict) and "client_id" not in token_info:
                main_config = json.loads(st.secrets["gcp_service_account"])
                app_info = main_config.get("web", main_config.get("installed", {}))
                token_info["client_id"] = app_info.get("client_id")
                token_info["client_secret"] = app_info.get("client_secret")
            
            creds = Credentials.from_authorized_user_info(token_info)
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
            return creds
        except Exception as e:
            st.error(f"⚠️ Token Error for {email}: {e}")
            return None
    return None
