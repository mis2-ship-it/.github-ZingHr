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

        if page.locator("input[id*='txtCompanyCode'], input[name*='CompanyCode']").is_visible():
            print("Entering Company Code...")
            page.locator("input[id*='txtCompanyCode'], input[name*='CompanyCode']").fill(ZING_COMPANY_CODE)

        print("Entering Employee Code / Username...")
        page.locator("input[id*='txtEmpCode'], input[id*='txtUserName'], input[name*='UserName']").first.fill(ZING_USERNAME)

        print("Entering Password...")
        pwd_field = page.locator("input[id*='txtPassword'], input[type='password']").first
        pwd_field.fill(ZING_PASSWORD)

        print("Submitting login credentials...")
        try:
            pwd_field.press("Enter")
        except Exception:
            page.locator("a[id*='Login'], a[id*='btn'], input[type='submit'], button[type='submit']").first.click(timeout=10000)

        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)
        print("Login complete.")

        # 2. Direct jump to Reports Gallery
        reports_view_url = "https://portal.zinghr.com/2015/Pages/ReportsGallery/ReportsView.aspx"
        print(f"Navigating directly to Reports Gallery: {reports_view_url}")
        page.goto(reports_view_url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)

        # 3. Locate the Active Working Frame (Checks if page uses an iframe)
        target_frame = page
        for frame in page.frames:
            try:
                if frame.locator("text='Reports Gallery'").is_visible() or frame.locator("text='Current Data'").is_visible():
                    target_frame = frame
                    print(f"Located active Reports Gallery frame: {frame.name or 'unnamed_iframe'}")
                    break
            except Exception:
                continue

        # 4. Ensure 'Current Data' tab is active
        print("Ensuring 'Current Data' tab is selected...")
        current_data_tab = target_frame.get_by_text("Current Data", exact=False).first
        if current_data_tab.is_visible():
            current_data_tab.click()
            page.wait_for_timeout(2000)

        # 5. Search or Directly Select 'Super Employee Master'
        print("Locating 'Super Employee Master'...")
        # Check if the search input is accessible
        search_inputs = target_frame.locator("input[type='text'], input[placeholder*='search' i], #txtSearch, input[id*='search' i]")
        if search_inputs.count() > 0:
            try:
                print("Typing 'super' into the search filter...")
                search_box = search_inputs.first
                search_box.fill("super")
                search_box.press("Enter")
                page.wait_for_timeout(2000)
            except Exception as e:
                print(f"Search box bypass: {e}")

        # If Employee MIS accordion is present and collapsed, click to expand
        emp_mis_header = target_frame.locator("text='Employee MIS'").first
        if emp_mis_header.is_visible():
            print("Expanding 'Employee MIS' accordion...")
            try:
                emp_mis_header.click()
                page.wait_for_timeout(1500)
            except Exception:
                pass

        # Click 'Super Employee Master'
        print("Clicking 'Super Employee Master'...")
        super_emp_item = target_frame.locator("text='Super Employee Master'").first
        super_emp_item.wait_for(state="visible", timeout=30000)
        super_emp_item.click()
        page.wait_for_timeout(4000)

        # 6. Expand 'Select time period & employees' if needed
        time_period_header = target_frame.locator("text='Select time period & employees'").first
        if time_period_header.is_visible():
            if not target_frame.locator("text='Show me reports for Status'").is_visible():
                time_period_header.click()
                page.wait_for_timeout(1500)

        # 7. Configure 'Show me reports for Status' -> Select All
        print("Checking Status dropdown...")
        status_dropdown = target_frame.locator("xpath=//span[contains(text(), 'Show me reports for Status') or contains(text(), 'Status')]/following::a[1] | //div[contains(., 'Status')]//button | //div[contains(., 'Status')]//a[contains(@class,'dropdown')]").first
        if status_dropdown.is_visible():
            status_text = status_dropdown.inner_text()
            if "all selected" not in status_text.lower():
                status_dropdown.click()
                page.wait_for_timeout(1000)
                select_all_chk = target_frame.locator("input[type='checkbox'][value*='multiselect-all'], label:has-text('Select all'), input[id*='chkAll']").first
                if select_all_chk.is_visible():
                    select_all_chk.check()
                    page.wait_for_timeout(1000)
                status_dropdown.click()
                page.wait_for_timeout(1000)

        # 8. Configure 'for Employee Code' -> Select All
        print("Checking Employee Code dropdown...")
        emp_code_dropdown = target_frame.locator("xpath=//span[contains(text(), 'for Employee Code') or contains(text(), 'Employee Code')]/following::a[1] | //div[contains(., 'Employee Code')]//button | //div[contains(., 'Employee Code')]//a[contains(@class,'dropdown')]").first
        if emp_code_dropdown.is_visible():
            emp_code_text = emp_code_dropdown.inner_text()
            if "all selected" not in emp_code_text.lower():
                emp_code_dropdown.click()
                page.wait_for_timeout(1000)
                select_all_emp = target_frame.locator("input[type='checkbox'][value*='multiselect-all'], label:has-text('Select all'), input[id*='chkAll']").first
                if select_all_emp.is_visible():
                    select_all_emp.check()
                    page.wait_for_timeout(1000)
                emp_code_dropdown.click()
                page.wait_for_timeout(1000)

        page.wait_for_timeout(2000)

        # 9. Click 'Export to Excel'
        print("Clicking 'Export to Excel' button...")
        export_btn = target_frame.locator("#btnExport, button:has-text('Export to Excel'), input[value*='Export to Excel'], a:has-text('Export to Excel')").first
        export_btn.wait_for(state="visible", timeout=30000)
        export_btn.click()
        page.wait_for_timeout(5000)
        print("Report generation requested successfully.")

        # 10. Switch to 'Processed Saved Reports (All)' tab
        print("Switching to 'Processed Saved Reports' queue...")
        processed_tab = target_frame.get_by_text("Processed Saved Reports", exact=False).first
        processed_tab.wait_for(state="visible", timeout=30000)
        processed_tab.click()
        page.wait_for_timeout(5000)

        # 11. Polling loop: Wait up to 7 minutes (420s) for the download arrow icon
        print("Waiting for report processing to complete (monitoring queue for up to 7 minutes)...")
        max_wait_seconds = 420
        poll_interval = 20
        elapsed = 0
        download_ready = False

        download_icon = target_frame.locator("table tbody tr:first-child a[title*='Download'], table tbody tr:first-child i[class*='download'], table tbody tr:first-child span[class*='download'], table tbody tr:first-child a:has(i), table tbody tr:first-child td:nth-child(5) a").first

        while elapsed < max_wait_seconds:
            if download_icon.is_visible():
                row_text = target_frame.locator("table tbody tr:first-child").inner_text()
                if "xlsx" in row_text.lower() and "error" not in row_text.lower():
                    print(f"Download icon is ready! (Elapsed: {elapsed} seconds)")
                    download_ready = True
                    break

            print(f"Still processing... ({elapsed}s / {max_wait_seconds}s). Refreshing Processed queue...")
            page.wait_for_timeout(poll_interval * 1000)
            elapsed += poll_interval

            try:
                processed_tab.click()
            except Exception:
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
                processed_tab = target_frame.get_by_text("Processed Saved Reports", exact=False).first
                processed_tab.click()

        if not download_ready:
            page.screenshot(path="downloads/timeout_queue.png")
            raise TimeoutError("Report did not finish processing within 7 minutes.")

        # 12. Trigger download from first row
        print("Triggering download from the first row...")
        with page.expect_download(timeout=180000) as download_info:
            download_icon.click()

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
