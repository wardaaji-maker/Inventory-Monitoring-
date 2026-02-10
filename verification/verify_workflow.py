import os
from playwright.sync_api import sync_playwright
import json

# Define the HTML file path
html_file_path = os.path.abspath("backend/Index.html")

# Mock Data
mock_clients = [
    {
        "row": 2,
        "clientId": "C001",
        "clientName": "Alice",
        "storeCode": "STORE_A",
        "pic": "Bob",
        "phone": "085712345678",
        "status": "LEADS",
        "progress": "PERKENALAN",
        "nextFollowUp": "2023-10-27T00:00:00.000Z",
        "overdueFollowUps": 1,
        "totalFollowUps": 1,
        "latestFeedback": "NO RESPON"
    }
]

# Mock Stats
mock_stats = {
    "totalClients": 1,
    "dueFollowups": 0,
    "overdueClients": 1,
    "followUpPercentage": "100.0",
    "progressCounts": {"PERKENALAN": 1},
    "feedbackCounts": {"NO RESPON": 1},
    "storeStats": {},
    "picStats": {}
}

def run_verification():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Inject Mock Script
        mock_script = f"""
        class Runner {{
            constructor() {{
                this._successCallback = null;
                this._failureCallback = null;
            }}
            withSuccessHandler(cb) {{
                this._successCallback = cb;
                return this;
            }}
            withFailureHandler(cb) {{
                this._failureCallback = cb;
                return this;
            }}
            getAllClients() {{
                if(this._successCallback) setTimeout(() => this._successCallback({json.dumps(mock_clients)}), 50);
            }}
            getDashboardStats() {{
                if(this._successCallback) setTimeout(() => this._successCallback({json.dumps(mock_stats)}), 50);
            }}
            getClientsNeedingFollowup() {{
                if(this._successCallback) setTimeout(() => this._successCallback([]), 50);
            }}
            getClientFollowupHistory(row) {{
                if(this._successCallback) setTimeout(() => this._successCallback([]), 50);
            }}
            recordFollowup(row, data) {{
                if(this._successCallback) setTimeout(() => this._successCallback(), 50);
            }}
        }}

        window.google = {{
            script: {{}}
        }};

        Object.defineProperty(window.google.script, 'run', {{
            get: function() {{
                return new Runner();
            }}
        }});
        """

        page.add_init_script(mock_script)
        page.goto(f"file://{html_file_path}")

        # Switch to Clients Tab
        page.click("button[onclick=\"showTab('clients')\"]")
        page.wait_for_selector("#clientsContent table")

        print("Data Loaded.")

        # 1. Verify Single WhatsApp -> Modal Trigger
        # Ensure modal is hidden initially
        if page.is_visible("#followupModal"):
            print("FAIL: Modal visible initially.")

        # Click Single WA Button
        # It opens a new window, we need to handle that or ignore it
        with page.expect_popup():
            page.click("#clientsContent button.btn-success")

        # Check if Modal Opened
        if page.is_visible("#followupModal"):
            print("PASS: Modal opened after clicking Single WA.")
            # Verify Pre-fill
            type_val = page.evaluate("document.getElementById('followupType').value")
            if type_val == "Message":
                print("PASS: Type pre-filled as Message.")
            else:
                print(f"FAIL: Type is {type_val}")
        else:
            print("FAIL: Modal did not open.")

        page.click("text=Cancel") # Close modal

        # 2. Verify Bulk WhatsApp -> Modal Trigger
        page.check("input[value='2']") # Select client
        page.click("#clientsTab button:has-text('Bulk WA')") # Open panel
        page.click("#clientsTab button:has-text('Generate Links')") # Open bulk modal

        page.wait_for_selector("#waModal")

        # Click "Send" in the list
        # It opens popup and SHOULD open followup modal (over the bulk modal)
        with page.expect_popup():
            page.click("#waList .btn-success")

        if page.is_visible("#followupModal"):
            print("PASS: Modal opened after clicking Bulk WA Send.")
        else:
            print("FAIL: Modal did not open after Bulk WA Send.")

        browser.close()

if __name__ == "__main__":
    run_verification()
