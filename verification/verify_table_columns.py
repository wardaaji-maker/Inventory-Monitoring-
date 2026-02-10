import os
import sys
from playwright.sync_api import sync_playwright

# Mock data
MOCK_CLIENTS = [
    {
        "row": 2,
        "clientId": "C001",
        "clientName": "Test Client 1",
        "storeCode": "S001",
        "phone": "123456789",
        "email": "test1@example.com",
        "address": "Address 1",
        "createdAt": "2023-01-01T00:00:00.000Z",
        "pic": "Agent Smith",
        "status": "LEADS",
        "progress": "PERKENALAN",
        "followUps": [],
        "nextFollowUp": "2023-02-01T00:00:00.000Z",
        "overdueFollowUps": 0,
        "totalFollowUps": 5,
        "latestFeedback": "Interested"
    },
    {
        "row": 3,
        "clientId": "C002",
        "clientName": "Test Client 2",
        "storeCode": "S002",
        "phone": "987654321",
        "email": "test2@example.com",
        "address": "Address 2",
        "createdAt": "2023-01-02T00:00:00.000Z",
        "pic": "Agent Doe",
        "status": "LEADS",
        "progress": "TERHUBUNG",
        "followUps": [],
        "nextFollowUp": None,
        "overdueFollowUps": 0,
        "totalFollowUps": 2,
        "latestFeedback": "Closed"
    }
]

def verify_table_columns():
    backend_path = os.path.abspath("backend/Index.html")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Load the page
        page.goto(f"file://{backend_path}")

        # Inject mock google.script.run
        page.evaluate("""
            window.google = {
                script: {
                    run: {
                        withSuccessHandler: function(callback) {
                            this.callback = callback;
                            return this;
                        },
                        withFailureHandler: function(callback) {
                            return this;
                        },
                        getAllClients: function() {
                            // Simulate async return
                            setTimeout(() => {
                                this.callback(JSON.parse(clientsData));
                            }, 100);
                        },
                        getDashboardStats: function() {
                             this.callback({
                                totalClients: 2,
                                dueFollowups: 0,
                                overdueClients: 0,
                                avgFollowUps: 3.5,
                                followUpPercentage: 100,
                                progressCounts: {},
                                feedbackCounts: {}
                             });
                        },
                        getClientsNeedingFollowup: function() {
                            // No-op for this test
                        }
                    }
                }
            };
        """)

        # Inject data
        page.evaluate(f"window.clientsData = '{import_json(MOCK_CLIENTS)}'")

        # Trigger load
        page.evaluate("init()")

        # Switch to Clients tab
        page.evaluate("showTab('clients')")

        # Wait for table to render
        page.wait_for_selector("#clientsContent table", timeout=5000)

        # Screenshot
        page.screenshot(path="verification/table_columns.png")

        # Verify Headers
        headers = page.locator("#clientsContent table thead th").all_inner_texts()
        print(f"Headers found: {headers}")

        if "PIC" not in headers:
            print("FAIL: 'PIC' column missing in header")
            sys.exit(1)

        if "Follow-ups" not in headers:
            print("FAIL: 'Follow-ups' column missing in header")
            sys.exit(1)

        # Verify Row Data
        rows = page.locator("#clientsContent table tbody tr")
        count = rows.count()
        print(f"Rows found: {count}")

        if count != 2:
            print(f"FAIL: Expected 2 rows, found {count}")
            sys.exit(1)

        # Check first row
        row1_texts = rows.nth(0).locator("td").all_inner_texts()
        print(f"Row 1 Data: {row1_texts}")

        # ID, Name, PIC, Phone, Status, Progress, Feedback, Next, Count, Actions
        # Index of PIC should be 2
        actual_pic = row1_texts[2]
        if actual_pic != "Agent Smith":
             print(f"FAIL: Row 1 PIC expected 'Agent Smith', got '{actual_pic}'")
             sys.exit(1)

        # Index of Count should be 8
        actual_count = row1_texts[8]
        # Note: inner_text might pick up styling, but usually just text.
        # The html is: <span ...>5</span>. Inner text should be '5'.
        if "5" not in actual_count:
             print(f"FAIL: Row 1 Follow-up Count expected '5', got '{actual_count}'")
             sys.exit(1)

        print("SUCCESS: Columns verified.")
        browser.close()

def import_json(data):
    import json
    return json.dumps(data)

if __name__ == "__main__":
    verify_table_columns()
