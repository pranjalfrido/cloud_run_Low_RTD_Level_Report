"""
Frido Low RTD Level Report — Email Sender (Microsoft Graph API)
Generates the Low RTD Excel report and emails it via Outlook/M365.

Sender : reports@myfrido.com
Usage  : python send_low_rtd_email.py
"""

import base64
import json
import os
from datetime import datetime
from pathlib import Path

import msal
import requests

from generate_low_rtd_report import main as generate_report

# ── Microsoft 365 OAuth credentials (loaded from environment variables) ───────
# Set these as Cloud Run Job environment variables — never hardcode secrets.
# Local run: export AZURE_CLIENT_ID=... AZURE_TENANT_ID=... AZURE_CLIENT_SECRET=...

CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SENDER_EMAIL  = "reports@myfrido.com"

TO_EMAILS = [
    "kunal.Kumbhar@myfrido.com",
    "navnath.k@myfrido.com",
    "gaurav.k@myfrido.com",
    "saurabh.m@myfrido.com",
    "sagar.k@myfrido.com",
]

CC_EMAILS = [
    "hemant.k@myfrido.com",
    "Ipshit.k@myfrido.com",
    "nikhil.c@myfrido.com",
    "pravin.k@myfrido.com",
    "amit.c@myfrido.com",
    "suraj.s@myfrido.com",
    "yash.s@myfrido.com",
    "yash.p@myfrido.com",
    "abhishek.v@myfrido.com",
    "tushar.g@myfrido.com",
    "pranjal.t@myfrido.com",
    "Syed.h@myfrido.com",
    "nikhil.k@myfrido.com",
    "Rahul.m@myfrido.com",
]

GRAPH_SCOPE    = ["https://graph.microsoft.com/.default"]
GRAPH_SEND_URL = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Failed to get token: {result.get('error_description', result)}")
    return result["access_token"]

# ── Email body ────────────────────────────────────────────────────────────────

def build_email_body(report_date_str, stats):
    return f"""
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222222; margin: 0; padding: 20px;">

<p>Hi All,</p>

<p>Please find attached the <strong>Low RTD Level Report for {report_date_str}</strong>.</p>

<p style="margin-top: 32px;">
  Thanks and Regards<br>
  <strong>Data Analytics Team</strong><br>
  Frido
</p>

</body>
</html>
""".strip()

# ── Send via Microsoft Graph ──────────────────────────────────────────────────

def send_email(report_path: str, stats: dict):
    report_date_str = datetime.now().strftime("%d %b %Y")
    subject         = f"Frido Low RTD Level Report - {report_date_str}"

    path = Path(report_path)
    with open(path, "rb") as f:
        attachment_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": build_email_body(report_date_str, stats),
            },
            "toRecipients": [
                {"emailAddress": {"address": email}} for email in TO_EMAILS
            ],
            "ccRecipients": [
                {"emailAddress": {"address": email}} for email in CC_EMAILS
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name":        path.name,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": attachment_b64,
                }
            ],
        },
        "saveToSentItems": True,
    }

    print("\nAcquiring Microsoft 365 token...")
    token = get_access_token()

    print(f"Sending email to: {', '.join(TO_EMAILS)}")
    print(f"CC: {', '.join(CC_EMAILS)}")
    response = requests.post(
        GRAPH_SEND_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        data=json.dumps(payload),
    )

    if response.status_code == 202:
        print("Email sent successfully!")
    else:
        raise RuntimeError(
            f"Graph API error {response.status_code}: {response.text}"
        )

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    report_path, stats = generate_report()

    if report_path is None:
        print("No report generated — email not sent.")
        return

    send_email(report_path, stats)

if __name__ == "__main__":
    main()
