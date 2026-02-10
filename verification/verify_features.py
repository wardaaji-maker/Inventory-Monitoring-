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
        "pic": "Bob",
        "phone": "08123456789",
        "status": "LEADS",
        "progress": "PERKENALAN",
        "nextFollowUp": "2023-10-27T00:00:00.000Z", # Past date
        "overdueFollowUps": 1,
        "totalFollowUps": 1
    },
    {
        "row": 3,
        "clientId": "C002",
        "clientName": "Charlie",
        "pic": "Dave",
        "phone": "+6281234567890",
        "status": "LEADS",
        "progress": "PENDEKATAN",
        "nextFollowUp": "2023-10-30T00:00:00.000Z", # Future
        "overdueFollowUps": 0,
        "totalFollowUps": 0
    },
    {
        "row": 4,
        "clientId": "C003",
        "clientName": "Eve",
        "pic": "Bob",
        "phone": "812345",
        "status": "LEADS",
        "progress": "TERHUBUNG",
        "nextFollowUp": None,
        "overdueFollowUps": 0,
        "totalFollowUps": 2
    }
]

mock_stats = {
    "totalClients": 3,
    "dueFollowups": 0,
    "overdueClients": 1,
    "followUpPercentage": "66.7",
    "progressCounts": {"PERKENALAN": 1, "PENDEKATAN": 1, "TERHUBUNG": 1},
    "feedbackCounts": {"No Feedback": 3},
    "storeStats": {},
    "picStats": {}
}

def run_verification():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Inject Mock Script
        mock_script = f"""
        window.google = {{
            script: {{
                run: {{
                    withSuccessHandler: function(callback) {{
                        this._successCallback = callback;
                        return this;
                    }},
                    withFailureHandler: function(callback) {{
                        this._failureCallback = callback;
                        return this;
                    }},
                    getAllClients: function() {{
                        setTimeout(() => this._successCallback({json.dumps(mock_clients)}), 100);
                    }},
                    getDashboardStats: function() {{
                        setTimeout(() => this._successCallback({json.dumps(mock_stats)}), 100);
                    }},
                    getClientsNeedingFollowup: function() {{
                         setTimeout(() => this._successCallback([]), 100);
                    }},
                    getClientFollowupHistory: function(row) {{
                         setTimeout(() => this._successCallback([]), 100);
                    }}
                }}
            }}
        }};
        """

        page.add_init_script(mock_script)
        page.goto(f"file://{html_file_path}")

        # Click All Clients tab to make it visible
        page.click("button[onclick=\"showTab('clients')\"]")

        # Wait for data load
        page.wait_for_selector("#clientsContent table", timeout=5000)

        print("Initial Data Loaded.")

        # 1. Verify Search
        print("Verifying Search...")
        page.fill("#searchClients", "Alice")
        page.dispatch_event("#searchClients", "keyup")
        # Wait for filter to apply (it's sync but DOM update takes a tick)
        page.wait_for_timeout(100)

        visible_rows = page.locator("#clientsContent tbody tr").count()
        if visible_rows == 1:
            print("PASS: Search for 'Alice' returned 1 row.")
        else:
            print(f"FAIL: Search for 'Alice' returned {visible_rows} rows.")

        page.fill("#searchClients", "") # Clear search
        page.dispatch_event("#searchClients", "keyup")
        page.wait_for_timeout(100)

        # 2. Verify PIC Filter
        print("Verifying PIC Filter...")
        page.select_option("#picFilterClients", "Bob")
        page.wait_for_timeout(100)
        visible_rows = page.locator("#clientsContent tbody tr").count()
        # Alice (Bob) and Eve (Bob) -> 2 rows
        if visible_rows == 2:
            print("PASS: PIC Filter 'Bob' returned 2 rows.")
        else:
            print(f"FAIL: PIC Filter 'Bob' returned {visible_rows} rows.")

        page.select_option("#picFilterClients", "") # Reset

        # 3. Verify WhatsApp Link Generation
        print("Verifying WhatsApp Logic...")
        # Check select box for Alice (Row 2)
        # Checkbox value is row ID
        page.check("input[value='2']")

        # Open Panel
        page.click("#clientsTab button:has-text('Bulk WA')")

        # Open Modal
        page.click("#clientsTab button:has-text('Generate Links')")
        page.wait_for_selector("#waModal", state="visible")

        # Check generated link
        # Alice phone: 08123456789 -> 628123456789
        expected_url_part = "wa.me/628123456789"

        links = page.locator("#waList a").all()
        if len(links) == 1:
            href = links[0].get_attribute("href")
            if expected_url_part in href:
                print(f"PASS: Generated WhatsApp link is correct: {href}")
            else:
                print(f"FAIL: Generated WhatsApp link incorrect. Got: {href}")
        else:
            print(f"FAIL: Expected 1 link, found {len(links)}")

        page.click("text=Done")

        # 4. Verify Mobile Responsiveness
        print("Verifying Mobile Layout...")
        page.set_viewport_size({"width": 375, "height": 667})
        # Check if sidebar buttons are stacked (display block/grid usually implies they have width)
        # We can check if the sidebar width is 100% or close to 375 (minus padding)
        sidebar_box = page.locator(".sidebar").bounding_box()
        if sidebar_box["width"] > 300: # 375 - padding
             print(f"PASS: Sidebar is full width on mobile ({sidebar_box['width']}px).")
        else:
             print(f"FAIL: Sidebar width is {sidebar_box['width']}px, expected full width.")

        # Check table overflow
        table_container_overflow = page.evaluate("document.querySelector('.table-container').style.overflowX || window.getComputedStyle(document.querySelector('.table-container')).overflowX")
        if table_container_overflow == "auto":
             print("PASS: Table container has overflow-x: auto.")
        else:
             print(f"FAIL: Table container overflow-x is {table_container_overflow}")

        # Check for Store and Follow-ups columns
        headers = page.locator("#clientsContent thead th").all_inner_texts()
        if "Store" in headers and "Follow-ups" in headers:
             print("PASS: 'Store' and 'Follow-ups' columns are present.")
        else:
             print(f"FAIL: Missing columns. Headers found: {headers}")

        browser.close()

if __name__ == "__main__":
    run_verification()
