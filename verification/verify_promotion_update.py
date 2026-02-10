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
        "content": "Monthly Sale", # Replaces Progress
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
    "contentCounts": {"Monthly Sale": 1},
    "feedbackCounts": {"NO RESPON": 1},
    "storeStats": {
        "STORE_A": {
            "total": 1,
            "content": {"Monthly Sale": 1},
            "feedback": {"NO RESPON": 1}
        }
    },
    "picStats": {}
}

mock_promotions = ["Monthly Sale", "New Arrival", "Special Offer"]

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
            getPromotionTypes() {{
                if(this._successCallback) setTimeout(() => this._successCallback({json.dumps(mock_promotions)}), 50);
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
        page.set_viewport_size({"width": 1280, "height": 800})
        page.goto(f"file://{html_file_path}")

        # Wait for data
        page.wait_for_selector(".stat-card")

        print("Data Loaded.")

        # 1. Verify Content Column in Table
        # Switch to Clients Tab
        page.click("button[onclick=\"showTab('clients')\"]")
        page.wait_for_selector("#clientsContent table")

        headers = page.locator("#clientsContent thead th").all_inner_texts()
        if "Content" in headers and "Progress" not in headers:
            print("PASS: Table Header updated to 'Content'.")
        else:
            print(f"FAIL: Table Header incorrect. Found: {headers}")

        # Verify Row Data
        row_text = page.locator("#clientsContent tbody tr").first.inner_text()
        if "Monthly Sale" in row_text:
            print("PASS: Row displays 'Monthly Sale'.")
        else:
            print("FAIL: Row missing content.")

        # 2. Verify Modal Dropdown
        # Click Follow-up Button
        page.click("#clientsContent button:has-text('Follow-up')") # Or the icon button
        # Actually my button has <i class="fas fa-plus"></i> inside.
        # Let's use selector:
        page.locator("#clientsContent button").nth(0).click()

        page.wait_for_selector("#followupModal")

        # Check Label
        label = page.locator("label:has-text('Content (Promotion)')")
        if label.count() > 0:
            print("PASS: Modal Label is 'Content (Promotion)'.")
        else:
            print("FAIL: Modal Label incorrect.")

        # Check Dropdown Options
        # Wait for options to populate (async)
        page.wait_for_timeout(100)
        options = page.locator("#followupContent option").all_inner_texts()
        if "New Arrival" in options:
            print("PASS: Modal Dropdown populated with dynamic options.")
        else:
            print(f"FAIL: Modal Dropdown missing options. Found: {options}")

        # 3. Verify Dashboard Breakdown
        # Switch to Dashboard
        page.click("button[onclick=\"showTab('dashboard')\"]")
        # Wait for breakdown
        page.wait_for_selector("#breakdownContent h4")

        content = page.locator("#breakdownContent").inner_text()
        if "Content Breakdown" in content:
            print("PASS: Dashboard shows 'Content Breakdown'.")
        else:
            print("FAIL: Dashboard missing 'Content Breakdown'.")

        browser.close()

if __name__ == "__main__":
    run_verification()
