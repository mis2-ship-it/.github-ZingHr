import os
import sys
import time
from playwright.sync_api import sync_playwright

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

def find_reports_frame(page):
    """Deeply inspects main page and all frames/iframes."""
    frames_to_check = page.frames
    log(f"Total frames detected on page: {len(frames_to_check)}")
    
    for i, frame in enumerate(frames_to_check):
        try:
            url = frame.url
            name = frame.name
            body_text = frame.evaluate("() => document.body ? document.body.innerText : ''")
            log(f"Frame #{i} [{name}] URL: {url[:60]}... Text length: {len(body_text)}")
            
            # Check for signatures of ZingHR Reports Gallery
            if any(k in body_text for k in ["Employee MIS", "Super Employee Master", "Current Data", "Reports Gallery", "Reports"]):
                log(f"==> MATCH FOUND in Frame #{i} ({name or url})")
                return frame
        except Exception as e:
            log(f"Frame #{i} evaluate note: {e}")
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
            viewport={"width": 1600, "height": 1000},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # ----------------------------------------------------
        # 1. LOGIN
        # ----------------------------------------------------
        log(f"Navigating to ZingHR login portal: {LOGIN_URL}")
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
        page.wait_for_timeout(8000)
        log(f"Login complete. Current URL: {page.url}")

        # ----------------------------------------------------
        # 2. NAVIGATE TO REPORTS GALLERY VIA PORTAL NAVIGATION
        # ----------------------------------------------------
        log(f"Login complete. Current URL: {page.url}")
        page.wait_for_timeout(5000)

        # Method A: Click Reports Gallery directly from the ZingNext navigation menu
        log("Navigating to Reports Gallery via portal interface...")
        navigated = False
        try:
            # Look for menu icon / Reports link in ZingNext portal
            reports_link = page.locator("a:has-text('Reports'), a:has-text('Report Gallery'), a:has-text('Reports Gallery'), [title*='Report' i], a[href*='ReportsView']").first
            if reports_link.is_visible():
                log("Found Reports menu item. Clicking...")
                # If it opens in a new tab/window, capture it
                with context.expect_page(timeout=10000) as new_page_info:
                    reports_link.click()
                new_page = new_page_info.value
                new_page.wait_for_load_state("domcontentloaded")
                page = new_page
                navigated = True
                log(f"Switched to Reports window: {page.url}")
        except Exception as e:
            log(f"Portal UI menu click note: {e}")

        # Method B: If direct UI click didn't trigger, navigate via relative or legacy bridge
        if not navigated or "AccessDenied" in page.url:
            log("Attempting authenticated session navigation...")
            # Navigate using the origin session rather than direct cold URL
            page.evaluate("""() => {
                const reportAnchor = Array.from(document.querySelectorAll('a, button, span, div')).find(el => {
                    const txt = (el.innerText || el.textContent || '').toLowerCase();
                    return txt.includes('report gallery') || txt.includes('reports gallery') || txt.includes('reports');
                });
                if (reportAnchor) {
                    (reportAnchor.closest('a') || reportAnchor).click();
                } else {
                    window.location.href = '/2015/Pages/ReportsGallery/ReportsView.aspx';
                }
            }""")
            page.wait_for_timeout(8000)

        log(f"Post-navigation URL: {page.url}")
        page.screenshot(path="downloads/reports_gallery_loaded.png")

        # Check if AccessDenied was returned
        if "AccessDenied" in page.url:
            # Method C: Launch via the top-level app switcher / 9-dots menu if present
            log("Access Denied on direct path. Attempting launch via App Switcher / Sidebar...")
            page.goto("https://zingnext.zinghr.com/portal", wait_until="domcontentloaded")
            page.wait_for_timeout(5000)
            
            # Click sidebar menu / app launcher
            page.locator(".menu-icon, .hamburger, [class*='sidebar'], [class*='menu'], .fa-bars").first.click()
            page.wait_for_timeout(2000)
            
            # Click Reports
            with context.expect_page(timeout=15000) as new_page_info:
                page.locator("text='Reports Gallery', text='Reports', a:has-text('Reports')").first.click()
            page = new_page_info.value
            page.wait_for_load_state("domcontentloaded")
            log(f"Successfully arrived via sidebar at: {page.url}")

        # ----------------------------------------------------
        # 3. IDENTIFY TARGET FRAME
        # ----------------------------------------------------
        target_frame = find_reports_frame(page)

        # ----------------------------------------------------
        # 4. CLICK 'CURRENT DATA' TAB IF NOT ACTIVE
        # ----------------------------------------------------
        log("Ensuring 'Current Data' tab is active...")
        target_frame.evaluate("""() => {
            const tabs = Array.from(document.querySelectorAll('a, span, li, button'));
            for (const t of tabs) {
                if (t.innerText && t.innerText.trim() === 'Current Data') {
                    t.click();
                    break;
                }
            }
        }""")
        page.wait_for_timeout(3000)

        # ----------------------------------------------------
        # 5. EXPAND 'EMPLOYEE MIS' & SELECT 'SUPER EMPLOYEE MASTER'
        # ----------------------------------------------------
        log("Expanding categories and selecting 'Super Employee Master'...")

        # Step A: Attempt to expand all tree nodes and search
        target_frame.evaluate("""() => {
            // Click any tree toggles / folders / plus icons
            document.querySelectorAll('.tree-toggle, .fa-plus, .fa-folder, .accordion-toggle, [data-toggle="collapse"]').forEach(el => el.click());

            // Type 'super' into the search input if present
            const inputs = Array.from(document.querySelectorAll("input[type='text'], input[placeholder*='search' i], #txtSearch, input[id*='search' i]"));
            for (const input of inputs) {
                if (input.offsetParent !== null) {
                    input.value = 'super';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }

            // Click search magnifying button
            const btn = document.querySelector(".input-group-addon, .input-group-btn, [id*='btnSearch'], button:has(.fa-search), a:has(.fa-search), .fa-search");
            if (btn) (btn.closest('button') || btn.closest('a') || btn).click();
        }""")
        page.wait_for_timeout(4000)

        # Step B: Click 'Super Employee Master' item
        log("Clicking 'Super Employee Master'...")
        clicked = target_frame.evaluate("""() => {
            const candidates = Array.from(document.querySelectorAll('a, li, span, div, p'));
            for (const el of candidates) {
                const txt = el.innerText ? el.innerText.trim() : '';
                // Match Super Employee Master, ignore 'With CTC'
                if (txt.includes('Super Employee Master') && !txt.includes('CTC')) {
                    const clickable = el.closest('a') || el.closest('li') || el;
                    clickable.scrollIntoView();
                    clickable.click();
                    
                    // If href has javascript, evaluate it
                    if (clickable.href && clickable.href.startsWith('javascript:')) {
                        const code = decodeURIComponent(clickable.href.replace('javascript:', ''));
                        try { eval(code); } catch(e) {}
                    }
                    return true;
                }
            }
            return false;
        }""")

        if not clicked:
            log("JS click did not trigger. Searching through all child frames for the element...")
            found_in_subframe = False
            for f in page.frames:
                try:
                    f_clicked = f.evaluate("""() => {
                        const candidates = Array.from(document.querySelectorAll('a, li, span'));
                        for (const el of candidates) {
                            const txt = el.innerText ? el.innerText.trim() : '';
                            if (txt.includes('Super Employee Master') && !txt.includes('CTC')) {
                                (el.closest('a') || el).click();
                                return true;
                            }
                        }
                        return false;
                    }""")
                    if f_clicked:
                        log(f"Clicked 'Super Employee Master' inside subframe: {f.name or f.url}")
                        target_frame = f
                        found_in_subframe = True
                        break
                except Exception:
                    continue

            if not found_in_subframe:
                page.screenshot(path="downloads/debug_before_timeout.png")
                raise RuntimeError("Could not find 'Super Employee Master' in any frame. Check downloads/debug_before_timeout.png")

        log("Super Employee Master selected. Waiting for report parameters to render...")
        page.wait_for_timeout(6000)

        # ----------------------------------------------------
        # 6. CONFIGURE FILTERS (SELECT ALL)
        # ----------------------------------------------------
        log("Configuring filters ('Select time period & employees')...")
        target_frame.evaluate("""() => {
            // Expand accordion
            document.querySelectorAll('a, div, span, h4, h5').forEach(el => {
                if (el.textContent && el.textContent.includes('Select time period & employees')) {
                    el.click();
                }
            });

            // Open all multiselect dropdowns that don't have all selected
            const triggers = Array.from(document.querySelectorAll('a.dropdown-toggle, button.multiselect, a[class*="multiselect"], div[class*="dropdown"] > a'));
            triggers.forEach(t => {
                const text = (t.textContent || t.innerText || '').toLowerCase();
                if (!text.includes('all selected')) {
                    t.click();
                }
            });

            // Check all select-all checkboxes
            document.querySelectorAll("input[type='checkbox']").forEach(chk => {
                const val = (chk.value || chk.id || chk.name || '').toLowerCase();
                const parent = (chk.parentElement ? chk.parentElement.textContent : '').toLowerCase();
                if (val.includes('multiselect-all') || val.includes('chkall') || parent.includes('select all')) {
                    if (!chk.checked) chk.click();
                }
            });

            // Close open dropdowns
            triggers.forEach(t => {
                if (t.getAttribute('aria-expanded') === 'true') {
                    t.click();
                }
            });
        }""")
        page.wait_for_timeout(3000)

        # ----------------------------------------------------
        # 7. TRIGGER 'EXPORT TO EXCEL'
        # ----------------------------------------------------
        log("Triggering 'Export to Excel' button...")
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

        log("Export requested. Waiting 6s before checking queue...")
        page.wait_for_timeout(6000)

        # ----------------------------------------------------
        # 8. SWITCH TO 'PROCESSED SAVED REPORTS (ALL)'
        # ----------------------------------------------------
        log("Switching to 'Processed Saved Reports (All)' tab...")
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

        # ----------------------------------------------------
        # 9. POLL QUEUE FOR DOWNLOAD READY (UP TO 7 MINUTES)
        # ----------------------------------------------------
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

        # ----------------------------------------------------
        # 10. DOWNLOAD FILE
        # ----------------------------------------------------
        log("Downloading file...")
        with page.expect_download(timeout=180000) as download_info:
            target_frame.evaluate("""() => {
                const firstRow = document.querySelector('table tbody tr:first-child');
                const dlLink = firstRow.querySelector('a[title*="Download" i], i.fa-download, i[class*="download"], a:has(i), td:nth-child(5) a');
                if (dlLink) dlLink.click();
            }""")

        download = download_info.value
        download.save_as(LOCAL_FILE_PATH)
        log(f"Super Employee Master successfully downloaded and saved to: {LOCAL_FILE_PATH}")
        browser.close()

def main():
    try:
        download_from_zinghr()
    except Exception as e:
        log(f"Execution Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
