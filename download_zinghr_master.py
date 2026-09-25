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
    print(f"Navigating to ZingHR portal login: {LOGIN_URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # 1. Login sequence
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)

        if page.is_visible("input[id*='txtCompanyCode'], input[name*='CompanyCode']"):
            print("Entering Company Code...")
            page.fill("input[id*='txtCompanyCode'], input[name*='CompanyCode']", ZING_COMPANY_CODE)

        print("Entering Employee Code / Username...")
        page.fill("input[id*='txtEmpCode'], input[id*='txtUserName'], input[name*='UserName']", ZING_USERNAME)

        print("Entering Password...")
        pwd_field = page.locator("input[id*='txtPassword'], input[type='password']").first
        pwd_field.fill(ZING_PASSWORD)

        print("Submitting login credentials...")
        try:
            pwd_field.press("Enter")
        except Exception:
            page.click("a[id*='Login'], a[id*='btn'], input[type='submit'], button[type='submit']", timeout=10000)

        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)
        print("Login complete.")

        # 2. Direct jump to Reports Gallery (Bypasses the 9-dots menu)
        reports_view_url = "https://portal.zinghr.com/2015/Pages/ReportsGallery/ReportsView.aspx"
        print(f"Navigating directly to Reports Gallery: {reports_view_url}")
        page.goto(reports_view_url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(4000)

        # 3. Ensure 'Current Data' tab is active & search/select 'Super Employee Master'
        print("Ensuring 'Current Data' tab is selected...")
        current_data_tab = "a:has-text('Current Data'), text='Current Data'"
        if page.is_visible(current_data_tab):
            page.click(current_data_tab)
            page.wait_for_timeout(2000)

        print("Selecting 'Super Employee Master' from left menu...")
        # Match the exact text block from the sidebar
        super_emp_item = "text='Super Employee Master' >> visible=true"
        page.wait_for_selector(super_emp_item, timeout=30000)
        page.click(super_emp_item)
        page.wait_for_timeout(3000)

        # 4. Trigger Report Generation ("Export to Excel" button)
        print("Clicking 'Export to Excel' button...")
        export_btn = "button:has-text('Export to Excel'), input[value*='Export to Excel'], a:has-text('Export to Excel'), #btnExport"
        page.wait_for_selector(export_btn, timeout=30000)
        page.click(export_btn)
        page.wait_for_timeout(5000)
        print("Report generation requested successfully.")

        # 5. Switch to 'Processed Saved Reports (All)' tab
        print("Switching to 'Processed Saved Reports' queue...")
        processed_tab = "a:has-text('Processed Saved Reports'), text='Processed Saved Reports'"
        page.wait_for_selector(processed_tab, timeout=30000)
        page.click(processed_tab)
        page.wait_for_timeout(4000)

        # 6. Polling loop: Wait up to 7 minutes (420s) for the download arrow icon to appear
        print("Waiting for report processing to complete (monitoring queue for up to 7 minutes)...")
        max_wait_seconds = 420
        poll_interval = 20
        elapsed = 0
        download_ready = False

        # Selector for the active download arrow in the first row
        download_icon_selector = "table tbody tr:first-child a[title*='Download'], table tbody tr:first-child i[class*='download'], table tbody tr:first-child span[class*='download'], table tbody tr:first-child a:has(i), table tbody tr:first-child td:nth-child(5) a"

        while elapsed < max_wait_seconds:
            # Check if the download icon is present and clickable in row 1
            if page.is_visible(download_icon_selector):
                # Verify that it is not still showing the spinner/loader
                row_text = page.locator("table tbody tr:first-child").inner_text()
                if "xlsx" in row_text.lower() and not ("error" in row_text.lower()):
                    print(f"Download icon is ready! (Elapsed: {elapsed} seconds)")
                    download_ready = True
                    break

            print(f"Still processing... ({elapsed}s / {max_wait_seconds}s). Refreshing Processed queue...")
            page.wait_for_timeout(poll_interval * 1000)
            elapsed += poll_interval

            # Re-click the tab or refresh the table view to poll the latest status
            try:
                page.click(processed_tab)
            except Exception:
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
                page.click(processed_tab)

        if not download_ready:
            # Take screenshot for diagnosis if it didn't finish in 7 mins
            page.screenshot(path="downloads/timeout_queue.png")
            raise TimeoutError("Report did not finish processing within 7 minutes.")

        # 7. Download the ready file
        print("Triggering download from the first row...")
        with page.expect_download(timeout=180000) as download_info:
            page.click(download_icon_selector)

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        print(f"File successfully downloaded and saved to: {LOCAL_FILE_PATH}")
        browser.close()

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
