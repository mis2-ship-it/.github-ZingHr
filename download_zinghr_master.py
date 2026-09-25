import os
import sys
import json
from datetime import datetime
from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- Environment Credentials ---
ZING_COMPANY_CODE = os.environ.get("ZING_COMPANY_CODE")
ZING_USERNAME = os.environ.get("ZING_USERNAME")
ZING_PASSWORD = os.environ.get("ZING_PASSWORD")
GCP_SA_JSON = os.environ.get("GCP_SA_KEY") # Service account JSON string from GitHub secrets

# Google Drive Target Folder ID
DRIVE_FOLDER_ID = "1Qil8HhgEiYf7GV1Oic8HOJoQ-13bz3aC"

DOWNLOAD_DIR = os.path.abspath("downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
LOCAL_FILE_PATH = os.path.join(DOWNLOAD_DIR, "Super_Employee_Master.xlsx")

def download_from_zinghr():
    print("Starting ZingHR download...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto("https://app.zinghr.com/login", wait_until="networkidle", timeout=60000)

        # Handle Company Code
        if page.is_visible("input#CompanyCode, input[name='CompanyCode']"):
            page.fill("input#CompanyCode, input[name='CompanyCode']", ZING_COMPANY_CODE)
            page.click("button:has-text('Proceed'), button:has-text('Next')")
            page.wait_for_timeout(2000)

        # Login
        page.fill("input#UserName, input[name='UserName'], input[type='text']", ZING_USERNAME)
        page.fill("input#Password, input[name='Password'], input[type='password']", ZING_PASSWORD)
        page.click("button[type='submit'], input[type='submit'], button:has-text('Sign In'), button:has-text('Login')")
        page.wait_for_load_state("networkidle", timeout=45000)

        # Navigate to Super Employee Master & Export
        page.click("text='Reports' >> visible=true")
        page.wait_for_timeout(1500)
        page.click("text='Employee Master' >> visible=true, text='Super Employee Master' >> visible=true")

        with page.expect_download(timeout=120000) as download_info:
            page.click("button:has-text('Export'), input[value*='Export'], a:has-text('Download Excel')")

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        browser.close()
        print(f"Downloaded to runner: {LOCAL_FILE_PATH}")

def upload_to_google_drive():
    print("Authenticating with Google Drive API...")
    sa_info = json.loads(GCP_SA_JSON)
    credentials = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    drive_service = build("drive", "v3", credentials=credentials)

    # Optional: Append date stamp to the file name, or keep it fixed
    today_str = datetime.now().strftime("%Y-%m-%d")
    drive_filename = f"Super_Employee_Master_{today_str}.xlsx"

    file_metadata = {
        "name": drive_filename,
        "parents": [DRIVE_FOLDER_ID]
    }
    media = MediaFileUpload(
        LOCAL_FILE_PATH,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    print(f"Uploading {drive_filename} to Google Drive folder 'ZingHR_Files'...")
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name"
    ).execute()

    print(f"Successfully uploaded: {uploaded_file.get('name')} (ID: {uploaded_file.get('id')})")

def main():
    if not all([ZING_COMPANY_CODE, ZING_USERNAME, ZING_PASSWORD, GCP_SA_JSON]):
        print("Error: Missing one or more required environment secrets.", file=sys.stderr)
        sys.exit(1)

    download_from_zinghr()
    upload_to_google_drive()

if __name__ == "__main__":
    main()
