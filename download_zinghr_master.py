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
        page.goto(reports_view_url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)

        # ----------------------------------------------------
        # 3. IDENTIFY WORKING FRAME (CRITICAL)
        # ----------------------------------------------------
        # Search across all frames for the frame containing the Reports DOM
        target_scope = page
        for frame in page.frames:
            try:
                has_content = frame.evaluate("""() => {
                    const text = document.body ? document.body.innerText : '';
                    return text.includes('Reports Gallery') || text.includes('Employee MIS') || text.includes('Current Data') || text.includes('Super Employee Master');
                }""")
                if has_content:
                    target_scope = frame
                    print(f"Located active Reports Gallery frame: {frame.name or frame.url}")
                    break
            except Exception:
                continue

        # ----------------------------------------------------
        # 4. SELECT 'SUPER EMPLOYEE MASTER' DIRECTLY IN FRAME
        # ----------------------------------------------------
        print("Selecting 'Super Employee Master' inside the identified frame...")

        # Step A: Filter by typing 'super' in the search box inside target_scope
        try:
            target_scope.evaluate("""() => {
                const inputs = Array.from(document.querySelectorAll("input[type='text'], input[placeholder*='search' i], #txtSearch"));
                for (const input of inputs) {
                    if (input.offsetParent !== null) {
                        input.value = 'super';
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        break;
                    }
                }
                const btn = document.querySelector(".input-group-addon, .input-group-btn, [id*='btnSearch'], button, a:has(.fa-search)");
                if (btn) btn.click();
            }""")
        except Exception as e:
            print(f"Search box evaluation note: {e}")

        page.wait_for_timeout(3000)

        # Step B: Click 'Super Employee Master' link directly via JS inside target_scope
        selected = target_scope.evaluate("""() => {
            // Find all anchor tags, list items, and clickable elements
            const candidates = Array.from(document.querySelectorAll('a, li, span, div'));
            for (const el of candidates) {
                const text = el.innerText ? el.innerText.trim() : '';
                // Must start with 'Super Employee Master' and NOT be 'With CTC'
                if (text.startsWith('Super Employee Master') && !text.includes('With CTC')) {
                    el.scrollIntoView();
                    // If element or its parent is an anchor with href="javascript:...", execute it
                    const anchor = el.tagName === 'A' ? el : el.closest('a');
                    if (anchor && anchor.href && anchor.href.startsWith('javascript:')) {
                        const jsCode = decodeURIComponent(anchor.href.replace('javascript:', ''));
                        try { eval(jsCode); } catch(err) { anchor.click(); }
                    } else {
                        el.click();
                    }
                    return true;
                }
            }
            return false;
        }""")

        if not selected:
            # Fallback: Find elements via partial text locator on target_scope (NOT page)
            print("Direct JS click did not locate item; trying locator on target_scope...")
            target_scope.locator("xpath=//a[contains(., 'Super Employee Master') and not(contains(., 'With CTC'))] | //li[contains(., 'Super Employee Master') and not(contains(., 'With CTC'))]").first.click(timeout=30000)

        print("Waiting for Super Employee Master view to load...")
        page.wait_for_timeout(5000)

        # ----------------------------------------------------
        # 5. EXPAND 'Select time period & employees' & SELECT ALL
        # ----------------------------------------------------
        print("Configuring filters ('Select time period & employees')...")
        target_scope.evaluate("""() => {
            // Expand accordion if collapsed
            document.querySelectorAll('a, div, span, h4, h5').forEach(el => {
                if (el.innerText && el.innerText.includes('Select time period & employees')) {
                    el.click();
                }
            });

            // Open all multi-select dropdowns that do not already show 'All selected'
            const triggers = Array.from(document.querySelectorAll('a.dropdown-toggle, button.multiselect, a[class*="multiselect"], div[class*="dropdown"] > a'));
            triggers.forEach(t => {
                if (t.innerText && !t.innerText.toLowerCase().includes('all selected')) {
                    t.click();
                }
            });

            // Check all select-all checkboxes
            document.querySelectorAll("input[type='checkbox']").forEach(chk => {
                const val = (chk.value || chk.id || chk.name || '').toLowerCase();
                const parent = (chk.parentElement ? chk.parentElement.innerText : '').toLowerCase();
                if (val.includes('multiselect-all') || val.includes('chkall') || parent.includes('select all')) {
                    if (!chk.checked) chk.click();
                }
            });

            // Close dropdowns
            triggers.forEach(t => {
                if (t.getAttribute('aria-expanded') === 'true') t.click();
            });
        }""")
        page.wait_for_timeout(3000)

        # ----------------------------------------------------
        # 6. TRIGGER 'EXPORT TO EXCEL'
        # ----------------------------------------------------
        print("Triggering 'Export to Excel'...")
        export_done = target_scope.evaluate("""() => {
            const btns = Array.from(document.querySelectorAll("input[type='submit'], input[type='button'], button, a"));
            for (const b of btns) {
                const val = (b.value || b.innerText || '').toLowerCase();
                if (val.includes('export to excel')) {
                    b.scrollIntoView();
                    b.click();
                    return true;
                }
            }
            const el = document.getElementById('btnExport');
            if (el) { el.click(); return true; }
            return false;
        }""")

        if not export_done:
            target_scope.locator("#btnExport, input[value*='Export to Excel'], button:has-text('Export to Excel'), a:has-text('Export to Excel')").first.click(timeout=20000)

        print("Export to Excel triggered successfully. Waiting 6s before checking queue...")
        page.wait_for_timeout(6000)

        # ----------------------------------------------------
        # 7. SWITCH TO 'PROCESSED SAVED REPORTS (ALL)'
        # ----------------------------------------------------
        print("Switching to 'Processed Saved Reports (All)' tab...")
        target_scope.evaluate("""() => {
            const tabs = Array.from(document.querySelectorAll('a, span, li, button'));
            for (const tab of tabs) {
                if (tab.innerText && tab.innerText.includes('Processed Saved Reports')) {
                    tab.scrollIntoView();
                    tab.click();
                    break;
                }
            }
        }""")
        page.wait_for_timeout(6000)

        # ----------------------------------------------------
        # 8. POLLING LOOP: WAIT FOR REPORT COMPLETION
        # ----------------------------------------------------
        print("Waiting for report processing to complete (monitoring queue for up to 7 minutes)...")
        max_wait_seconds = 420
        poll_interval = 20
        elapsed = 0
        download_ready = False

        while elapsed < max_wait_seconds:
            ready_status = target_scope.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                if (!firstRow) return false;
                const text = firstRow.innerText.toLowerCase();
                if (text.includes('error')) return 'error';

                const dl = firstRow.querySelector('a[title*="Download"], i.fa-download, i[class*="download"], a:has(i)');
                if (dl && text.includes('xlsx')) return true;
                return false;
            }""")

            if ready_status == 'error':
                raise RuntimeError("ZingHR server encountered an error processing the report.")

            if ready_status is True:
                print(f"Download icon is ready! (Elapsed: {elapsed} seconds)")
                download_ready = True
                break

            print(f"Still processing in ZingHR queue ({elapsed}s / {max_wait_seconds}s)...")
            page.wait_for_timeout(poll_interval * 1000)
            elapsed += poll_interval

            # Click tab to refresh queue
            target_scope.evaluate("""() => {
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
        # 9. TRIGGER FILE DOWNLOAD
        # ----------------------------------------------------
        print("Initiating file download from row 1...")
        with page.expect_download(timeout=180000) as download_info:
            target_scope.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                const dl = firstRow.querySelector('a[title*="Download"], i.fa-download, i[class*="download"], a:has(i), td:nth-child(5) a');
                if (dl) dl.click();
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
