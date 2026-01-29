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
        "phone": "08123456789",
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
        "phone": "+6281234567890",
        "status": "LEADS",
        "progress": "PENDEKATAN",
        "nextFollowUp": "2023-10-30T00:00:00.000Z",
        "overdueFollowUps": 0,
        "totalFollowUps": 0,
        "latestFeedback": ""
    }
]

# Mock Stats with specific breakdown structure
mock_stats = {
    "totalClients": 2,
    "dueFollowups": 0,
    "overdueClients": 1,
    "followUpPercentage": "50.0",
    "progressCounts": {"PERKENALAN": 1, "PENDEKATAN": 1},
    "feedbackCounts": {"NO RESPON": 1, "No Feedback": 1},
    "storeStats": {
        "STORE_A": {
            "total": 1, "contribution": "100.0", "dueToday": 0, "overdue": 1,
            "progress": {"PERKENALAN": 1},
            "feedback": {"NO RESPON": 1}
        },
        "STORE_B": {
            "total": 1, "contribution": "0.0", "dueToday": 0, "overdue": 0,
            "progress": {"PENDEKATAN": 1},
            "feedback": {"No Feedback": 1}
        }
    },
    "picStats": {
        "Bob": {
            "total": 1, "contribution": "100.0", "dueToday": 0, "overdue": 1,
            "progress": {"PERKENALAN": 1},
            "feedback": {"NO RESPON": 1}
        },
        "Dave": {
            "total": 1, "contribution": "0.0", "dueToday": 0, "overdue": 0,
            "progress": {"PENDEKATAN": 1},
            "feedback": {"No Feedback": 1}
        }
    }
}

def run_verification():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Inject Mock Script with correct chaining handling using a getter
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

        # Wait for dashboard content
        # Note: #statsGrid is empty initially and filled by renderDashboard.
        # If it timeouts, it means renderDashboard wasn't called or failed.
        # We wait for the .stat-card which is injected into statsGrid.
        page.wait_for_selector(".stat-card")

        # Wait for breakdown content explicitly
        page.wait_for_selector("#breakdownContent table")
        print("Dashboard Loaded.")

        # 1. Verify Breakdown Tables Existence
        # By default "By Store" is selected
        content = page.locator("#breakdownContent").inner_text()

        if "Overview" in content and "Progress Breakdown" in content and "Feedback Breakdown" in content:
            print("PASS: Breakdown sections (Overview, Progress, Feedback) are present.")
        else:
            print(f"FAIL: Missing breakdown sections. Content: {content[:100]}...")

        # 2. Verify Table Data (Store A / PERKENALAN)
        if "STORE_A" in content and "PERKENALAN" in content:
             print("PASS: STORE_A and PERKENALAN labels visible in breakdown.")

        # 3. Verify Filter Logic
        # Switch to Clients Tab
        page.click("button[onclick=\"showTab('clients')\"]")
        page.wait_for_selector("#clientsContent table")

        print("Testing Filter...")

        # Test Case 1: Filter by PIC "Bob"
        page.select_option("#picFilterClients", "Bob")
        page.wait_for_timeout(100)
        rows = page.locator("#clientsContent tbody tr").count()
        if rows == 1:
             print("PASS: Filter by PIC 'Bob' works.")
        else:
             print(f"FAIL: Filter by PIC 'Bob' failed. Rows: {rows}")

        # Test Case 2: Filter by PIC with case sensitivity check (Mock "Bob", filter "BOB" if option allowed it, but options are populated from data)
        # We can't easily test mismatched case unless we inject mismatched option.
        # But we can verify "Dave"
        page.select_option("#picFilterClients", "Dave")
        page.wait_for_timeout(100)
        rows = page.locator("#clientsContent tbody tr").count()
        if rows == 1:
             print("PASS: Filter by PIC 'Dave' works.")

        page.select_option("#picFilterClients", "") # Reset

        # Test Case 3: Search by Name "Alice"
        page.fill("#searchClients", "Alice")
        page.dispatch_event("#searchClients", "keyup")
        page.wait_for_timeout(100)
        rows = page.locator("#clientsContent tbody tr").count()
        if rows == 1:
             print("PASS: Filter by Name 'Alice' works.")
        else:
             print(f"FAIL: Filter by Name 'Alice' failed. Rows: {rows}")

        # Test Case 4: Empty Search
        page.fill("#searchClients", "")
        page.dispatch_event("#searchClients", "keyup")
        page.wait_for_timeout(100)
        rows = page.locator("#clientsContent tbody tr").count()
        if rows == 2:
             print("PASS: Clearing search restores all rows.")

        browser.close()

if __name__ == "__main__":
    run_verification()
