import re
from googleapiclient.discovery import build

def get_jd_html(creds, doc_id):
    drive_service = build('drive', 'v3', credentials=creds)
    html_content = drive_service.files().export(fileId=doc_id, mimeType='text/html').execute()
    decoded_html = html_content.decode('utf-8')
    
    # Cleaning regex
    decoded_html = re.sub(r'margin-top:\s*[\d\.]+(pt|px|cm|in);?', 'margin-top: 0 !important;', decoded_html)
    decoded_html = re.sub(r'margin-bottom:\s*[\d\.]+(pt|px|cm|in);?', 'margin-bottom: 0 !important;', decoded_html)
    decoded_html = re.sub(r'padding-top:\s*[\d\.]+(pt|px|cm|in);?', 'padding-top: 0 !important;', decoded_html)
    decoded_html = re.sub(r'padding-bottom:\s*[\d\.]+(pt|px|cm|in);?', 'padding-bottom: 0 !important;', decoded_html)
    decoded_html = re.sub(r'\.c\d+\s*{[^}]+}', '', decoded_html)
    decoded_html = re.sub(r'(<(h[1-6]|p)[^>]*>)', r'\1', decoded_html, count=1)
    
    if "<body" in decoded_html:
        decoded_html = re.sub(r'<body[^>]*>', '<body style="margin:0; padding:0; background-color:#ffffff;">', decoded_html)

    style_fix = """
    <style>
        body { margin: 0 !important; padding: 0 !important; }
        body > *, body > div > * { margin-top: 0 !important; padding-top: 0 !important; }
        body, td, p, h1, h2, h3 { font-family: Arial, Helvetica, sans-serif !important; color: #000000 !important; }
        p { margin-bottom: 8px !important; margin-top: 0 !important; }
        ul, ol { margin-top: 0 !important; margin-bottom: 8px !important; padding-left: 25px !important; }
        li { margin-bottom: 2px !important; }
        li p { display: inline !important; margin: 0 !important; }
        h1, h2, h3 { margin-bottom: 10px !important; margin-top: 15px !important; }
        h1:first-child, h2:first-child, h3:first-child { margin-top: 0 !important; }
    </style>
    """
    
    clean_html = style_fix + decoded_html
    docs_service = build('docs', 'v1', credentials=creds)
    doc = docs_service.documents().get(documentId=doc_id).execute()
    return doc.get('title'), clean_html
