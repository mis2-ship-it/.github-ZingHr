import os
import sys
import time
import glob
from playwright.sync_api import sync_playwright

# Retrieve credentials securely from environment variables
ZING_COMPANY_CODE = os.environ.get("RIPL")  # e.g., ROYALOAK
ZING_USERNAME = os.environ.get("CON001")
ZING_PASSWORD = os.environ.get("Ravimalla@123")

DOWNLOAD_DIR = os.path.abspath("downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def run():
    if not all([ZING_COMPANY_CODE, ZING_USERNAME, ZING_PASSWORD]):
        print("Missing credentials in environment variables.", file=sys.stderr)
        sys.exit(1)

    with sync_playwright() as p:
        # Launch browser with download capabilities enabled
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Navigating to ZingHR login portal...")
        page.goto("https://app.zinghr.com/login", wait_until="networkidle", timeout=60000)

        # Fill Company Code / Domain if prompted on landing
        if page.is_visible("input#CompanyCode, input[name='CompanyCode']"):
            page.fill("input#CompanyCode, input[name='CompanyCode']", ZING_COMPANY_CODE)
            page.click("button:has-text('Proceed'), button:has-text('Next')")
            page.wait_for_timeout(2000)

        # Input User Credentials
        print("Entering credentials...")
        page.fill("input#UserName, input[name='UserName'], input[type='text']", ZING_USERNAME)
        page.fill("input#Password, input[name='Password'], input[type='password']", ZING_PASSWORD)
        page.click("button[type='submit'], input[type='submit'], button:has-text('Sign In'), button:has-text('Login')")

        page.wait_for_load_state("networkidle", timeout=45000)
        print("Successfully authenticated.")

        # --- Navigate to Reports / Super Employee Master ---
        # Adjust URL or click pathway to match your exact ZingHR menu layout
        print("Locating Super Employee Master report...")
        
        # Scenario A: Direct URL endpoint (if available in your tenant)
        # page.goto("https://app.zinghr.com/Report/SuperEmployeeMaster", wait_until="networkidle")

        # Scenario B: Menu Navigation
        page.click("text='Reports' >> visible=true")
        page.wait_for_timeout(1500)
        page.click("text='Employee Master' >> visible=true, text='Super Employee Master' >> visible=true")

        # Listen for download event upon clicking Export / Generate
        print("Triggering export download...")
        with page.expect_download(timeout=120000) as download_info:
            # Matches Export button (Excel / CSV)
            page.click("button:has-text('Export'), input[value*='Export'], a:has-text('Download Excel')")

        download = download_info.value
        suggested_name = download.suggested_filename
        final_path = os.path.join(DOWNLOAD_DIR, "Super_Employee_Master.xlsx")
        download.save_as(final_path)

        print(f"File successfully downloaded and saved to: {final_path}")
        browser.close()

if __name__ == "__main__":
    run()
