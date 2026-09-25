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

        # ----------------------------------------------------
        # 1. LOGIN
        # ----------------------------------------------------
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

        # ----------------------------------------------------
        # 2. OPEN REPORTS GALLERY & WAIT FOR HYDRATION
        # ----------------------------------------------------
        reports_view_url = "https://portal.zinghr.com/2015/Pages/ReportsGallery/ReportsView.aspx"
        print(f"Navigating directly to Reports Gallery: {reports_view_url}")
        page.goto(reports_view_url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(5000)

        # ----------------------------------------------------
        # 3. LOCATE ACTIVE FRAME (MAIN PAGE OR EMBEDDED IFRAME)
        # ----------------------------------------------------
        active_scope = page
        for frame in page.frames:
            try:
                has_reports = frame.evaluate("() => document.body.innerText.includes('Reports Gallery') || document.body.innerText.includes('Super Employee Master') || document.body.innerText.includes('Current Data')")
                if has_reports:
                    active_scope = frame
                    print(f"Target frame identified: {frame.name or 'embedded_frame'}")
                    break
            except Exception:
                continue

        # ----------------------------------------------------
        # 4. FILTER AND CLICK 'SUPER EMPLOYEE MASTER'
        # ----------------------------------------------------
        print("Filtering and selecting 'Super Employee Master'...")
        # Step A: Filter by typing 'super' in the left search box
        active_scope.evaluate("""() => {
            const inputs = Array.from(document.querySelectorAll("input[type='text'], input[placeholder*='search' i], #txtSearch"));
            for (const input of inputs) {
                if (input.offsetParent !== null) { // visible input
                    input.value = 'super';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
            const btn = document.querySelector(".input-group-addon, .input-group-btn, [id*='btnSearch'], button, a:has(.fa-search)");
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(3000)

        # Step B: Click 'Super Employee Master' item
        clicked_item = active_scope.evaluate("""() => {
            const items = Array.from(document.querySelectorAll('a, li, div, span'));
            for (const item of items) {
                const txt = item.innerText ? item.innerText.trim() : '';
                if (txt.startsWith('Super Employee Master') && !txt.includes('With CTC')) {
                    item.scrollIntoView();
                    item.click();
                    return true;
                }
            }
            return false;
        }""")

        if not clicked_item:
            # Fallback direct locator click
            page.locator("text='Super Employee Master'").first.click(timeout=15000)

        print("Waiting for ASP.NET AJAX UpdatePanel to load Super Employee Master view...")
        # CRITICAL: Wait until the right panel header updates to 'Super Employee Master'
        page.wait_for_selector("xpath=//h1[contains(., 'Super Employee Master')] | //h2[contains(., 'Super Employee Master')] | //h3[contains(., 'Super Employee Master')] | //div[contains(@class,'title') and contains(., 'Super Employee Master')] | //span[contains(., 'Super Employee Master')]", timeout=30000)
        page.wait_for_timeout(4000)

        # ----------------------------------------------------
        # 5. CONFIGURE FILTERS ('Select time period & employees')
        # ----------------------------------------------------
        print("Checking and selecting dropdown filters...")
        active_scope.evaluate("""() => {
            // Expand accordion if needed
            document.querySelectorAll('a, div, span').forEach(el => {
                if (el.innerText && el.innerText.includes('Select time period & employees')) {
                    el.click();
                }
            });

            // Open any multi-select dropdowns that are not 'All selected'
            const triggers = Array.from(document.querySelectorAll('a.dropdown-toggle, button.multiselect, a[class*="multiselect"], div[class*="dropdown"] > a'));
            triggers.forEach(t => {
                if (t.innerText && !t.innerText.toLowerCase().includes('all selected')) {
                    t.click();
                }
            });

            // Check all 'multiselect-all' or select-all checkboxes
            document.querySelectorAll("input[type='checkbox']").forEach(chk => {
                const val = (chk.value || chk.id || chk.name || '').toLowerCase();
                const parent = (chk.parentElement ? chk.parentElement.innerText : '').toLowerCase();
                if (val.includes('multiselect-all') || val.includes('chkall') || parent.includes('select all')) {
                    if (!chk.checked) chk.click();
                }
            });

            // Close triggers
            triggers.forEach(t => {
                if (t.getAttribute('aria-expanded') === 'true') t.click();
            });
        }""")
        page.wait_for_timeout(3000)

        # ----------------------------------------------------
        # 6. TRIGGER 'EXPORT TO EXCEL'
        # ----------------------------------------------------
        print("Locating and clicking 'Export to Excel' button...")
        # First ensure it's visible in the DOM
        page.wait_for_selector("xpath=//input[contains(@value, 'Export to Excel')] | //button[contains(., 'Export to Excel')] | //a[contains(., 'Export to Excel')]", timeout=30000)
        
        # Click via JavaScript directly to bypass any click-interception or overlays
        active_scope.evaluate("""() => {
            const targets = Array.from(document.querySelectorAll("input[type='submit'], input[type='button'], button, a"));
            for (const el of targets) {
                const text = (el.value || el.innerText || '').toLowerCase();
                if (text.includes('export to excel')) {
                    el.scrollIntoView();
                    el.click();
                    return true;
                }
            }
            return false;
        }""")
        print("Export to Excel successfully clicked. Waiting 6s for generation request...")
        page.wait_for_timeout(6000)

        # ----------------------------------------------------
        # 7. SWITCH TO 'PROCESSED SAVED REPORTS (ALL)'
        # ----------------------------------------------------
        print("Switching to 'Processed Saved Reports (All)' tab...")
        active_scope.evaluate("""() => {
            const tabs = Array.from(document.querySelectorAll('a, span, li, button'));
            for (const tab of tabs) {
                if (tab.innerText && tab.innerText.includes('Processed Saved Reports')) {
                    tab.scrollIntoView();
                    tab.click();
                    break;
                }
            }
        }""")
        page.wait_for_timeout(10000)

        # ----------------------------------------------------
        # 8. POLLING LOOP: WAIT 5-6 MINUTES FOR DOWNLOAD ARROW ICON
        # ----------------------------------------------------
        print("Waiting for report generation to complete in the queue (polling up to 7 minutes)...")
        max_wait_seconds = 420  # 7 minutes
        poll_interval = 20      # re-check every 20 seconds
        elapsed = 0
        download_ready = False

        while elapsed < max_wait_seconds:
            is_ready = active_scope.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                if (!firstRow) return false;

                const text = firstRow.innerText.toLowerCase();
                // If it contains error, abort
                if (text.includes('error')) return 'error';

                // Look for the download link/icon () in the 5th column or within the row
                const downloadLink = firstRow.querySelector('a[title*="Download"], i.fa-download, i[class*="download"], a:has(i)');
                if (downloadLink && text.includes('xlsx')) {
                    return true;
                }
                return false;
            }""")

            if is_ready == 'error':
                raise RuntimeError("ZingHR server reported an error processing this report.")

            if is_ready is True:
                print(f"Download icon is ready! (Elapsed: {elapsed} seconds)")
                download_ready = True
                break

            print(f"Report is still processing in ZingHR queue ({elapsed}s / {max_wait_seconds}s). Re-checking in {poll_interval}s...")
            page.wait_for_timeout(poll_interval * 1000)
            elapsed += poll_interval

            # Click 'Processed Saved Reports' tab again to refresh the table status
            active_scope.evaluate("""() => {
                const tabs = Array.from(document.querySelectorAll('a, span, li'));
                for (const t of tabs) {
                    if (t.innerText && t.innerText.includes('Processed Saved Reports')) {
                        t.click();
                        break;
                    }
                }
            }""")

        if not download_ready:
            page.screenshot(path="downloads/timeout_queue.png")
            raise TimeoutError("Report did not finish processing within 7 minutes.")

        # ----------------------------------------------------
        # 9. TRIGGER DOWNLOAD FROM FIRST ROW
        # ----------------------------------------------------
        print("Initiating file download from row 1...")
        with page.expect_download(timeout=180000) as download_info:
            active_scope.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                const downloadLink = firstRow.querySelector('a[title*="Download"], i.fa-download, i[class*="download"], a:has(i), td:nth-child(5) a');
                if (downloadLink) {
                    downloadLink.click();
                }
            }""")

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        print(f"Super Employee Master successfully downloaded and saved to: {LOCAL_FILE_PATH}")
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
