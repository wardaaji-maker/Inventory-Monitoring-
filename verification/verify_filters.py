from playwright.sync_api import sync_playwright
import os
import json

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Set viewport to see everything
    page.set_viewport_size({"width": 1280, "height": 800})

    # Mock Data
    mock_clients = [
        {
            "row": 2, "clientId": "C001", "clientName": "Client One", "phone": "123",
            "status": "LEADS", "progress": "PERKENALAN", "pic": "Sales1",
            "nextFollowUp": "2023-12-01", "overdueFollowUps": 0, "totalFollowUps": 1,
            "latestFeedback": "Good"
        },
        {
            "row": 3, "clientId": "C002", "clientName": "Client Two", "phone": "456",
            "status": "LEADS", "progress": "TERHUBUNG", "pic": "Sales2",
            "nextFollowUp": "2023-12-05", "overdueFollowUps": 0, "totalFollowUps": 2,
            "latestFeedback": "Connected"
        },
        {
            "row": 4, "clientId": "C003", "clientName": "Client Three", "phone": "789",
            "status": "LEADS", "progress": "PENDEKATAN", "pic": "Sales1",
            "nextFollowUp": "2023-11-20", "overdueFollowUps": 1, "totalFollowUps": 1,
            "latestFeedback": "Meeting set"
        }
    ]

    mock_dashboard = {
        "totalClients": 3, "dueFollowups": 1, "overdueClients": 1, "avgFollowUps": 1.3,
        "statusCounts": {"LEADS": 3}
    }

    # Inject Mock google.script.run
    # We serialize the data to JSON to safely inject it into the JS string
    clients_json = json.dumps(mock_clients)
    dashboard_json = json.dumps(mock_dashboard)

    page.add_init_script(f"""
        window.google = {{
            script: {{
                run: {{
                    withSuccessHandler: function(callback) {{
                        this._callback = callback;
                        return this;
                    }},
                    withFailureHandler: function(callback) {{
                        return this;
                    }},
                    getAllClients: function() {{
                        const data = {clients_json};
                        setTimeout(() => this._callback(data), 100);
                    }},
                    getDashboardStats: function() {{
                        const data = {dashboard_json};
                        setTimeout(() => this._callback(data), 100);
                    }},
                    getClientsNeedingFollowup: function() {{
                        setTimeout(() => this._callback([]), 100);
                    }},
                    getClientFollowupHistory: function(row) {{
                         setTimeout(() => this._callback([]), 100);
                    }}
                }}
            }}
        }};
    """)

    # Load the page
    cwd = os.getcwd()
    file_url = f"file://{cwd}/frontend/index.html"
    print(f"Loading {file_url}")
    page.goto(file_url)

    # Click the "All Clients" tab to make the table visible
    print("Clicking 'All Clients' tab...")
    page.click("button[onclick=\"showTab('clients')\"]")

    # Wait for table
    print("Waiting for table...")
    page.wait_for_selector("#clientsContent table", state="visible")

    # Check initial rows
    rows = page.locator("#clientsContent tbody tr")
    print(f"Initial rows: {rows.count()}")

    # Apply PIC Filter
    print("Filtering by PIC: Sales1")
    page.select_option("#picFilter", "Sales1")
    page.wait_for_timeout(500) # Wait for filter logic

    # Verify
    visible_rows_count = page.locator("#clientsContent tbody tr").count() # DOM elements remain but might be hidden?
    # Actually the filter function in JS re-renders the table: `renderTable(filtered, ...)`
    # So the rows in DOM should strictly equal the filtered result.
    rows_count = page.locator("#clientsContent tbody tr").count()
    print(f"Rows after PIC filter: {rows_count}")

    if rows_count != 2:
        print("FAILED: Expected 2 rows after PIC filter")

    # Apply Progress Filter
    print("Filtering by Progress: PERKENALAN")
    page.select_option("#progressFilter", "PERKENALAN")
    page.wait_for_timeout(500)

    rows_count = page.locator("#clientsContent tbody tr").count()
    print(f"Rows after Progress filter: {rows_count}")

    if rows_count != 1:
        print("FAILED: Expected 1 row after Progress filter")

    # Take Screenshot
    page.screenshot(path="verification/verification_filters.png")
    print("Screenshot saved to verification/verification_filters.png")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
