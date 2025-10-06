from flask import Flask, request, jsonify
from msal import ConfidentialClientApplication
import requests

app = Flask(__name__)

TENANT_ID = "4395a262-82b5-4485-b45b-bfbf365f5bc3"
CLIENT_ID = "d7391559-e3de-4734-b295-d48ecb903cd8"
CLIENT_SECRET = "52a2f3e-cccc-42da-9abc-1e4202b8dd67"

# Replace with your SharePoint details
SITE_HOSTNAME = "tascoutsourcing-my.sharepoint.com"
SITE_PATH = "/sites/YourSiteName"   # e.g. "/sites/AutomationTeam"
FILE_PATH = "https://tascoutsourcing-my.sharepoint.com/personal/automations_tascoutsourcing_com/Documents/Daily_reports/Business_Travel.xlsx"
WORKSHEET = "Sheet1"
TABLE = "TravelReports"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_token():
    app_ = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app_.acquire_token_for_client(SCOPES)
    return result["access_token"]

@app.route("/submit", methods=["POST"])
def submit():
    try:
        token = get_token()
        form_data = request.json
        row_values = [list(form_data.values())]

        # Step 1: Get siteId
        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}",
            headers={"Authorization": f"Bearer {token}"}
        )
        site_id = site_resp.json()["id"]

        # Step 2: Get driveItem id for Excel
        file_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{FILE_PATH}",
            headers={"Authorization": f"Bearer {token}"}
        )
        file_id = file_resp.json()["id"]

        # Step 3: Add row
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/workbook/worksheets/{WORKSHEET}/tables/{TABLE}/rows/add"
        resp = requests.post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"values": row_values}
        )

        return jsonify(resp.json())
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(port=3000, debug=True)
