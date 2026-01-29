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
        "phone": "085712345678", # Case: Starts with 0
        "status": "LEADS",
        "progress": "PERKENALAN",
        "nextFollowUp": "2023-10-27T00:00:00.000Z",
        "overdueFollowUps": 1,
        "totalFollowUps": 1,
        "latestFeedback": "NO RESPON"
    },
    {
        "row": 3,
        "clientId": "C002",
        "clientName": "Charlie",
        "storeCode": "STORE_B",
        "pic": "Dave",
        "phone": "85787654321", # Case: Starts with 8
        "status": "LEADS",
        "progress": "PENDEKATAN",
        "nextFollowUp": "2023-10-30T00:00:00.000Z",
        "overdueFollowUps": 0,
        "totalFollowUps": 0,
        "latestFeedback": ""
    }
]

# Mock Stats
mock_stats = {
    "totalClients": 2,
    "dueFollowups": 0,
    "overdueClients": 1,
    "followUpPercentage": "50.0",
    "progressCounts": {"PERKENALAN": 1, "PENDEKATAN": 1},
    "feedbackCounts": {"NO RESPON": 1, "No Feedback": 1},
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

        # 1. Verify WhatsApp Link Normalization
        # Open Modal for both users
        page.check("input[value='2']") # Alice (0857...)
        page.check("input[value='3']") # Charlie (857...)

        # Click "Bulk WA" to show panel
        page.click("#clientsTab button:has-text('Bulk WA')")

        page.click("#clientsTab button:has-text('Generate Links')")
        page.wait_for_selector("#waModal")

        links = page.locator("#waList a").all()

        found_alice = False
        found_charlie = False

        for link in links:
            href = link.get_attribute("href")
            # Alice: 0857... -> 62857...
            if "6285712345678" in href:
                found_alice = True
            # Charlie: 857... -> 62857...
            if "6285787654321" in href:
                found_charlie = True

        if found_alice:
            print("PASS: WhatsApp Link for 0857... -> 62857... correct.")
        else:
            print("FAIL: WhatsApp Link for Alice incorrect.")

        if found_charlie:
            print("PASS: WhatsApp Link for 857... -> 62857... correct.")
        else:
            print("FAIL: WhatsApp Link for Charlie incorrect.")

        page.click("text=Done")

        # 2. Verify Mobile View
        print("Testing Mobile View...")
        page.set_viewport_size({"width": 375, "height": 667})

        # Check if table is hidden
        table_visible = page.is_visible(".data-table")
        if not table_visible:
            print("PASS: Table is hidden on mobile.")
        else:
            print("FAIL: Table is visible on mobile.")

        # Check if cards are visible
        cards_visible = page.is_visible(".mobile-card")
        if cards_visible:
            print("PASS: Mobile Cards are visible.")
        else:
            print("FAIL: Mobile Cards are hidden.")

        browser.close()

if __name__ == "__main__":
    run_verification()
