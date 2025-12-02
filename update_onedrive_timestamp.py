import os
import requests
from datetime import datetime
from urllib.parse import quote

# ---------- CONFIG VIA VARIABLES D'ENV ----------
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
USER_PRINCIPAL_NAME = os.environ["USER_PRINCIPAL_NAME"]
ONEDRIVE_FILE_PATH = os.environ.get("ONEDRIVE_FILE_PATH", "Documents/Test Clément.xlsx")
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "Feuil1")
# ----------------------------------------

# 1. Authentification
token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
token_data = {
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials"
}

token_response = requests.post(token_url, data=token_data)
token_json = token_response.json()
token = token_json.get("access_token")

session = requests.Session()
session.headers.update({"Authorization": f"Bearer {token}"})

# 2. URL du workbook
encoded_path = quote(ONEDRIVE_FILE_PATH)
base_url = (
    "https://graph.microsoft.com/v1.0/"
    f"users/{USER_PRINCIPAL_NAME}/drive/root:/{encoded_path}:/workbook"
)
worksheet_url = f"{base_url}/worksheets('{WORKSHEET_NAME}')"

# 3. Récupération du usedRange
used = session.get(f"{worksheet_url}/usedRange(valuesOnly=true)").json()
values = used.get("values", [])
next_row = len(values) + 1 if values else 1

# 4. Écriture du timestamp
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
address = f"A{next_row}"

payload = {"values": [[timestamp]]}

print("Worksheet URL:", worksheet_url)
print("Payload:", payload)

url = f"{worksheet_url}/range(address='{address}')"
resp = session.patch(url, json=payload)

print("Status:", resp.status_code)