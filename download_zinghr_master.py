import os
import sys
import json
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- Environment Credentials ---
ZING_COMPANY_CODE = os.environ.get("ZING_COMPANY_CODE")
ZING_USERNAME = os.environ.get("ZING_USERNAME")
ZING_PASSWORD = os.environ.get("ZING_PASSWORD")
GCP_SA_JSON = os.environ.get("GCP_SA_KEY")

# Target Google Drive Folder: "ZingHR_Files"
DRIVE_FOLDER_ID = "1Qil8HhgEiYf7GV1Oic8HOJoQ-13bz3aC"

DOWNLOAD_DIR = os.path.abspath("downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
LOCAL_FILE_PATH = os.path.join(DOWNLOAD_DIR, "Super_Employee_Master.xlsx")

# Correct ZingHR classic enterprise login endpoint
LOGIN_URL = "https://portal.zinghr.com/2015/pages/authentication/login.aspx"

def download_from_zinghr():
    print(f"Navigating to ZingHR classic portal: {LOGIN_URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # Use domcontentloaded for ASP.NET pages to prevent networkidle hang
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)

        # 1. Company Code
        if page.is_visible("input[id*='txtCompanyCode'], input[name*='CompanyCode']"):
            print("Entering Company Code...")
            page.fill("input[id*='txtCompanyCode'], input[name*='CompanyCode']", ZING_COMPANY_CODE)

        # 2. Employee Code / Username
        print("Entering Employee Code / Username...")
        page.fill("input[id*='txtEmpCode'], input[id*='txtUserName'], input[name*='UserName']", ZING_USERNAME)

        # 3. Password
        print("Entering Password...")
        pwd_field = page.locator("input[id*='txtPassword'], input[type='password']").first
        pwd_field.fill(ZING_PASSWORD)

        # 4. Submit (Press Enter directly in the password field)
        print("Submitting login credentials via Enter key...")
        try:
            pwd_field.press("Enter")
        except Exception:
            # Fallback to common ZingHR login link/button element variants
            page.click("a[id*='Login'], a[id*='btn'], input[type='submit'], button[type='submit'], .login-btn", timeout=10000)
        
        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)
        print("Authentication processed. Navigating to Reports...")

        # Navigate to Reports / Super Employee Master
        page.click("text='Reports' >> visible=true")
        page.wait_for_timeout(2000)
        page.click("text='Super Employee Master' >> visible=true, text='Employee Master' >> visible=true")
        page.wait_for_timeout(3000)

        # Trigger Excel download
        print("Triggering download...")
        with page.expect_download(timeout=120000) as download_info:
            page.click("input[id*='btnExport'], input[value*='Export'], button:has-text('Export'), a:has-text('Download Excel')")

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        browser.close()
        print(f"File downloaded successfully to: {LOCAL_FILE_PATH}")

def upload_to_google_drive():
    print("Connecting to Google Drive API...")
    sa_info = json.loads(GCP_SA_JSON)
    credentials = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    drive_service = build("drive", "v3", credentials=credentials)

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

    print(f"Uploading {drive_filename} to Drive folder 'ZingHR_Files'...")
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name"
    ).execute()

    print(f"Uploaded successfully: {uploaded_file.get('name')} (ID: {uploaded_file.get('id')})")

def main():
    if not all([ZING_COMPANY_CODE, ZING_USERNAME, ZING_PASSWORD, GCP_SA_JSON]):
        print("Error: Missing one or more required secrets in GitHub environment.", file=sys.stderr)
        sys.exit(1)

    download_from_zinghr()
    upload_to_google_drive()

if __name__ == "__main__":
    main()
