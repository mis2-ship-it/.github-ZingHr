import os
import sys
import time
from playwright.sync_api import sync_playwright

# Force standard output to flush immediately on GitHub Actions
def log(msg):
    print(msg, flush=True)

ZING_COMPANY_CODE = os.environ.get("ZING_COMPANY_CODE", "")
ZING_USERNAME = os.environ.get("ZING_USERNAME", "")
ZING_PASSWORD = os.environ.get("ZING_PASSWORD", "")

LOGIN_URL = "https://portal.zinghr.com/2015/pages/authentication/login.aspx"
REPORTS_VIEW_URL = "https://portal.zinghr.com/2015/Pages/ReportsGallery/ReportsView.aspx"
LOCAL_FILE_PATH = os.environ.get("LOCAL_FILE_PATH", "downloads/Super_Employee_Master.xlsx")


def ensure_output_dir(file_path):
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)


def get_active_reports_frame(page):
    for frame in page.frames:
        try:
            has_reports = frame.evaluate("""() => {
                const text = document.body ? document.body.innerText : '';
                return text.includes('Reports Gallery') || 
                       text.includes('Current Data') || 
                       text.includes('Employee MIS') ||
                       text.includes('Super Employee Master');
            }""")
            if has_reports:
                log(f"Target working frame identified: {frame.name or frame.url}")
                return frame
        except Exception:
            continue
    return page.main_frame


def download_from_zinghr():
    log("==================================================")
    log("Starting ZingHR Super Employee Master Automation")
    log("==================================================")
    
    if not ZING_USERNAME or not ZING_PASSWORD:
        log("ERROR: ZING_USERNAME or ZING_PASSWORD secret is missing or empty!")
        sys.exit(1)

    ensure_output_dir(LOCAL_FILE_PATH)
    log(f"Navigating to ZingHR login portal: {LOGIN_URL}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ]
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # 1. Login sequence
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)

        if page.locator("input[id*='txtCompanyCode'], input[name*='CompanyCode']").is_visible():
            log("Entering Company Code...")
            page.locator("input[id*='txtCompanyCode'], input[name*='CompanyCode']").fill(ZING_COMPANY_CODE)

        log("Entering Employee Code / Username...")
        page.locator("input[id*='txtEmpCode'], input[id*='txtUserName'], input[name*='UserName']").first.fill(ZING_USERNAME)

        log("Entering Password...")
        pwd_field = page.locator("input[id*='txtPassword'], input[type='password']").first
        pwd_field.fill(ZING_PASSWORD)

        log("Submitting login credentials...")
        try:
            pwd_field.press("Enter")
        except Exception:
            page.locator("a[id*='Login'], a[id*='btn'], input[type='submit'], button[type='submit']").first.click(timeout=10000)

        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)
        log("Login complete.")

        # 2. Navigate to Reports Gallery
        log(f"Navigating to Reports Gallery: {REPORTS_VIEW_URL}")
        page.goto(REPORTS_VIEW_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(6000)

        target_frame = get_active_reports_frame(page)

        # 3. Ensure 'Current Data' tab
        log("Ensuring 'Current Data' tab is selected...")
        try:
            target_frame.evaluate("""() => {
                const tabs = Array.from(document.querySelectorAll('a, span, li'));
                for (const t of tabs) {
                    if (t.innerText && t.innerText.includes('Current Data')) {
                        t.click();
                        break;
                    }
                }
            }""")
        except Exception as e:
            log(f"Tab note: {e}")
        page.wait_for_timeout(3000)

        # 4. Search and select 'Super Employee Master'
        log("Searching for 'Super Employee Master'...")
        target_frame.evaluate("""() => {
            const inputs = Array.from(document.querySelectorAll("input[type='text'], input[placeholder*='search' i], #txtSearch"));
            for (const input of inputs) {
                if (input.offsetWidth > 0 && input.offsetHeight > 0) {
                    input.value = 'super';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
            const searchBtn = document.querySelector(".input-group-addon, .input-group-btn, [id*='btnSearch'], button:has(.fa-search), a:has(.fa-search), .fa-search");
            if (searchBtn) {
                (searchBtn.closest('button') || searchBtn.closest('a') || searchBtn).click();
            }

            document.querySelectorAll('a, li, div, span, h4, h5').forEach(el => {
                if (el.textContent && el.textContent.includes('Employee MIS')) {
                    el.click();
                }
            });
        }""")
        page.wait_for_timeout(3000)

        selected = target_frame.evaluate("""() => {
            const items = Array.from(document.querySelectorAll('a, li, span, div'));
            for (const item of items) {
                const txt = item.textContent ? item.textContent.trim() : '';
                if (txt.includes('Super Employee Master') && !txt.includes('With CTC')) {
                    const target = item.closest('a') || item.closest('li') || item;
                    target.scrollIntoView();
                    target.click();
                    return true;
                }
            }
            return false;
        }""")

        if not selected:
            log("Using fallback XPath locator for report...")
            target_frame.locator("xpath=//a[contains(., 'Super Employee Master') and not(contains(., 'With CTC'))] | //li[contains(., 'Super Employee Master') and not(contains(., 'With CTC'))]").first.click(timeout=20000)

        log("Waiting for report parameters to hydrate...")
        page.wait_for_timeout(5000)

        # 5. Configure dropdowns (Status = ALL, Emp Code = ALL)
        log("Configuring filters to 'All selected'...")
        target_frame.evaluate("""() => {
            document.querySelectorAll('a, div, span, h4, h5').forEach(el => {
                if (el.textContent && el.textContent.includes('Select time period & employees')) {
                    el.click();
                }
            });

            const triggers = Array.from(document.querySelectorAll('a.dropdown-toggle, button.multiselect, a[class*="multiselect"], div[class*="dropdown"] > a'));
            triggers.forEach(t => {
                const text = (t.textContent || t.innerText || '').toLowerCase();
                if (!text.includes('all selected')) {
                    t.click();
                }
            });

            document.querySelectorAll("input[type='checkbox']").forEach(chk => {
                const val = (chk.value || chk.id || chk.name || '').toLowerCase();
                const parent = (chk.parentElement ? chk.parentElement.textContent : '').toLowerCase();
                if (val.includes('multiselect-all') || val.includes('chkall') || parent.includes('select all')) {
                    if (!chk.checked) chk.click();
                }
            });

            triggers.forEach(t => {
                if (t.getAttribute('aria-expanded') === 'true') {
                    t.click();
                }
            });
        }""")
        page.wait_for_timeout(3000)

        # 6. Click 'Export to Excel'
        log("Triggering 'Export to Excel'...")
        exported = target_frame.evaluate("""() => {
            const btns = Array.from(document.querySelectorAll("input[type='submit'], input[type='button'], button, a"));
            for (const b of btns) {
                const val = (b.value || b.textContent || '').toLowerCase();
                if (val.includes('export to excel')) {
                    b.scrollIntoView();
                    b.click();
                    return true;
                }
            }
            const el = document.getElementById('btnExport');
            if (el) {
                el.scrollIntoView();
                el.click();
                return true;
            }
            return false;
        }""")

        if not exported:
            target_frame.locator("xpath=//input[contains(@value, 'Export to Excel')] | //button[contains(., 'Export to Excel')] | //a[contains(., 'Export to Excel')] | #btnExport").first.click(timeout=20000)

        log("Export requested. Switching to Processed tab...")
        page.wait_for_timeout(6000)

        # 7. Switch to 'Processed Saved Reports (All)'
        target_frame.evaluate("""() => {
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

        # 8. Polling loop
        log("Polling queue for download readiness (up to 7 mins)...")
        max_wait_seconds = 420
        poll_interval = 20
        elapsed = 0
        download_ready = False

        while elapsed < max_wait_seconds:
            poll_status = target_frame.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                if (!firstRow) return { ready: false, reason: 'no_row' };

                const text = firstRow.innerText.toLowerCase();
                if (text.includes('error')) return { ready: false, reason: 'server_error', text: text };

                const dlLink = firstRow.querySelector('a[title*="Download" i], i.fa-download, i[class*="download"], a:has(i), td:nth-child(5) a');
                const hasSpinner = firstRow.querySelector('.fa-spin, .spinner, img[src*="load"], [class*="spin"]');

                if (dlLink && !hasSpinner && text.includes('xlsx')) {
                    return { ready: true };
                }
                return { ready: false, reason: 'in_progress' };
            }""")

            if poll_status.get("reason") == "server_error":
                raise RuntimeError(f"ZingHR queue error: {poll_status.get('text')}")

            if poll_status.get("ready") is True:
                log(f"Download icon ready! Waited: {elapsed}s")
                download_ready = True
                break

            log(f"Processing ({elapsed}s / {max_wait_seconds}s)... waiting {poll_interval}s")
            page.wait_for_timeout(poll_interval * 1000)
            elapsed += poll_interval

            try:
                target_frame.evaluate("""() => {
                    const tabs = Array.from(document.querySelectorAll('a, span, li'));
                    for (const t of tabs) {
                        if (t.innerText && t.innerText.includes('Processed Saved Reports')) {
                            t.click();
                            break;
                        }
                    }
                }""")
            except Exception:
                pass

        if not download_ready:
            page.screenshot(path="downloads/timeout_queue.png")
            raise TimeoutError("Report processing timed out after 7 minutes.")

        # 9. Download
        log("Downloading file...")
        with page.expect_download(timeout=180000) as download_info:
            target_frame.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                const dlLink = firstRow.querySelector('a[title*="Download" i], i.fa-download, i[class*="download"], a:has(i), td:nth-child(5) a');
                if (dlLink) dlLink.click();
            }""")

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        log(f"Saved successfully to: {LOCAL_FILE_PATH}")
        browser.close()


def main():
    try:
        download_from_zinghr()
    except Exception as e:
        log(f"Execution Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
