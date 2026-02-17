from datetime import datetime, timezone, timedelta
from googleapiclient.discovery import build

def get_full_sheet_data(creds, sheet_id, sheet_name):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{sheet_name}!A:C").execute()
        return res.get('values', []) if res.get('values') else []
    except: return []

def get_send_log(creds, sheet_id):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        res = sheets.spreadsheets().values().get(spreadsheetId=sheet_id, range="SendLog!A:B").execute()
        values = res.get('values', [])
        return set((row[0].strip(), row[1].strip()) for row in values if len(row) >= 2)
    except Exception as e:
        print(f"Error reading SendLog: {e}")
        return set()

def append_to_send_log(creds, sheet_id, target_email, subject, sender_account):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        timestamp = datetime.now(ist_tz).strftime('%Y-%m-%d %I:%M:%S %p IST')
        
        body = {'values': [[target_email, subject, sender_account, timestamp]]}
        sheets.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range="SendLog!A:D",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()
    except Exception as e:
        print(f"Error appending to SendLog: {e}")
