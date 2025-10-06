import msal, requests, time, json
from msal import ConfidentialClientApplication
CLIENT_ID = "d7391559-e3de-4734-b295-d48ecb903cd8"
TENANT = "4395a262-82b5-4485-b45b-bfbf365f5bc3"  # e.g. contoso.onmicrosoft.com or tenant-id or 'common'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"
CLIENT_SECRET = '5U88Q~bFb0pBxQb6diOeNJXXvrs941bKasKaybOw' 
# SCOPES = ["Files.ReadWrite.All", "Sites.ReadWrite.All", "offline_access", "User.Read"]
SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f'https://login.microsoftonline.com/{TENANT}',
    client_credential=CLIENT_SECRET,
)

# Acquire a token for the client
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
access_token = token_response.get('access_token')
print(access_token)

import urllib.request, json, base64
import urllib.request, json

drive_id = "b!xAlDY9XeHk-46h02bfed_zE35ulWClpDseJA87Z7hnqYdkFwZ-4CQoxES2vMxemC"
item_id = "01VY5D56F6Y2GOVW7725BZO354PWSELRRZ"

url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/createSession"
body = b'{"persistChanges": true}'

req = urllib.request.Request(url, data=body, method="POST")
req.add_header("Authorization", "Bearer " + access_token)
req.add_header("Content-Type", "application/json")

try:
    with urllib.request.urlopen(req) as resp:
        data = json.load(resp)
        print(data)
except urllib.error.HTTPError as e:
    print("Status:", e.code)
    print("Reason:", e.reason)
    print("Response body:", e.read().decode())
