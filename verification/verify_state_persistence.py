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
        context = browser.new_context()
        page = context.new_page()

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

        # 1. Verify Tab Persistence (Session)
        print("Testing Tab Persistence...")
        page.click("button[onclick=\"showTab('clients')\"]")
        page.wait_for_selector("#clientsContent table")

        # Check Local Storage
        state = page.evaluate("localStorage.getItem('leadAppState')")
        if state and 'clients' in state:
            print("PASS: State saved to localStorage.")
        else:
            print(f"FAIL: State not saved. {state}")

        # 2. Verify Filter Persistence (Ajax Refresh)
        print("Testing Filter Persistence after Reload...")
        # Set Filter
        page.fill("#searchClients", "Alice")
        page.dispatch_event("#searchClients", "keyup")
        page.wait_for_timeout(100)

        # Trigger reload (mimic saveFollowup success)
        page.evaluate("loadClients()")
        page.wait_for_timeout(200) # Wait for mock data

        # Check if filter input still has value
        val = page.input_value("#searchClients")
        if val == "Alice":
            print("PASS: Search input retained value.")
        else:
            print(f"FAIL: Search input reset to '{val}'.")

        # Check if table is still filtered
        rows = page.locator("#clientsContent tbody tr").count()
        if rows == 1:
            print("PASS: Table is still filtered.")
        else:
            print(f"FAIL: Table reset/unfiltered. Rows: {rows}")

        # 3. Verify Refresh Persistence (Simulated Reload)
        print("Testing Browser Refresh Persistence...")
        page.reload()
        # Page reload clears the mock script injection in Playwright unless we use add_init_script again?
        # add_init_script persists for context.

        page.wait_for_selector("#clientsContent table")

        # Check if we are on 'clients' tab (active class)
        cls = page.get_attribute("button[onclick=\"showTab('clients')\"]", "class")
        if "active" in cls:
            print("PASS: Clients tab is active after reload.")
        else:
            print(f"FAIL: Clients tab not active. Class: {cls}")

        # Check input value restored
        val = page.input_value("#searchClients")
        if val == "Alice":
            print("PASS: Search input restored after reload.")
        else:
            print(f"FAIL: Search input not restored. Val: '{val}'")

        browser.close()

if __name__ == "__main__":
    run_verification()
