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
        # 4. SELECT 'SUPER EMPLOYEE MASTER' VIA DIRECT DOM CLICK
        # ----------------------------------------------------
        print("Selecting 'Super Employee Master' using direct DOM execution...")
        selected = active_scope.evaluate("""() => {
            // First, expand all collapsed accordions/sections
            document.querySelectorAll('.accordion, .panel-heading, [data-toggle="collapse"], a, div').forEach(el => {
                if (el.innerText && el.innerText.includes('Employee MIS')) {
                    el.click();
                }
            });

            // Find any element whose direct text contains 'Super Employee Master' (avoiding 'With CTC')
            const allElements = Array.from(document.querySelectorAll('a, li, div, span, p, h4, h5'));
            for (const el of allElements) {
                const text = el.innerText ? el.innerText.trim() : '';
                if (text.startsWith('Super Employee Master') && !text.includes('With CTC')) {
                    el.scrollIntoView();
                    el.click();
                    return true;
                }
            }
            return false;
        }""")

        # Fallback if DOM search didn't find it: type 'super' in search box and click magnifying glass
        if not selected:
            print("Direct click not triggered; using search box filter...")
            active_scope.evaluate("""() => {
                const inputs = Array.from(document.querySelectorAll("input[type='text'], input[placeholder*='search' i], #txtSearch"));
                if (inputs.length > 0) {
                    inputs[0].value = 'super';
                    inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
                    inputs[0].dispatchEvent(new Event('change', { bubbles: true }));
                }
                const btn = document.querySelector(".input-group-addon, .input-group-btn, [id*='btnSearch'], button, a:has(.fa-search)");
                if (btn) btn.click();
            }""")
            page.wait_for_timeout(3000)

            # Retry clicking 'Super Employee Master'
            active_scope.evaluate("""() => {
                const els = Array.from(document.querySelectorAll('a, li, div, span'));
                for (const el of els) {
                    if (el.innerText && el.innerText.trim().startsWith('Super Employee Master') && !el.innerText.includes('CTC')) {
                        el.click();
                        break;
                    }
                }
            }""")

        page.wait_for_timeout(4000)
        print("Super Employee Master selected.")

        # ----------------------------------------------------
        # 5. EXPAND OPTIONS & CONFIGURE DROPDOWNS (SELECT ALL)
        # ----------------------------------------------------
        print("Configuring 'Select time period & employees' filters...")
        active_scope.evaluate("""() => {
            // 1. Expand accordion if collapsed
            const accordions = Array.from(document.querySelectorAll('a, div, span, h4, h5'));
            for (const acc of accordions) {
                if (acc.innerText && acc.innerText.includes('Select time period & employees')) {
                    acc.click();
                    break;
                }
            }
        }""")
        page.wait_for_timeout(2000)

        # Set Status and Employee Code multi-selects to 'All'
        active_scope.evaluate("""() => {
            // Find dropdown buttons/links for Status & Employee Code
            const dropTriggers = Array.from(document.querySelectorAll('a.dropdown-toggle, button.multiselect, a[class*="multiselect"], div[class*="multiselect"]'));
            dropTriggers.forEach(trigger => {
                if (!trigger.innerText.toLowerCase().includes('all selected')) {
                    trigger.click();
                }
            });

            // Check all 'Select all' checkboxes
            const chks = Array.from(document.querySelectorAll('input[type="checkbox"][value*="multiselect-all"], input[type="checkbox"][id*="chkAll"], .multiselect-all input'));
            chks.forEach(chk => {
                if (!chk.checked) {
                    chk.click();
                }
            });

            // Close open dropdown popups
            dropTriggers.forEach(trigger => trigger.click());
        }""")
        page.wait_for_timeout(2000)

        # ----------------------------------------------------
        # 6. CLICK 'EXPORT TO EXCEL'
        # ----------------------------------------------------
        print("Triggering 'Export to Excel'...")
        export_clicked = active_scope.evaluate("""() => {
            const buttons = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"], a'));
            for (const b of buttons) {
                const val = (b.value || b.innerText || '').toLowerCase();
                if (val.includes('export to excel')) {
                    b.click();
                    return true;
                }
            }
            const exportById = document.getElementById('btnExport');
            if (exportById) {
                exportById.click();
                return true;
            }
            return false;
        }""")

        if not export_clicked:
            active_scope.locator("#btnExport, button:has-text('Export to Excel'), a:has-text('Export to Excel')").first.click(timeout=15000)

        print("Export to Excel triggered. Waiting 5s before switching to Processed tab...")
        page.wait_for_timeout(5000)

        # ----------------------------------------------------
        # 7. SWITCH TO 'PROCESSED SAVED REPORTS (ALL)' TAB
        # ----------------------------------------------------
        print("Navigating to 'Processed Saved Reports (All)'...")
        active_scope.evaluate("""() => {
            const tabs = Array.from(document.querySelectorAll('a, span, li, button'));
            for (const tab of tabs) {
                if (tab.innerText && tab.innerText.includes('Processed Saved Reports')) {
                    tab.click();
                    break;
                }
            }
        }""")
        page.wait_for_timeout(5000)

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
