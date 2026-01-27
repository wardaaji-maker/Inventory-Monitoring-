import os
import sys
import json
from playwright.sync_api import sync_playwright

# Mock Stats Data
MOCK_STATS = {
    "totalClients": 10,
    "dueFollowups": 2,
    "overdueClients": 1,
    "avgFollowUps": 2.5,
    "followUpPercentage": 80.0,
    "progressCounts": {"PERKENALAN": 5, "PENDEKATAN": 3, "TERHUBUNG": 2},
    "feedbackCounts": {"Interested": 5, "Closed": 2, "No Response": 3},
    "storeStats": {
        "STORE_A": {
            "total": 5,
            "followedUp": 4,
            "contribution": "80.0",
            "dueToday": 1,
            "overdue": 0,
            "progress": {"PERKENALAN": 3, "PENDEKATAN": 2},
            "feedback": {"Interested": 4, "No Response": 1}
        },
        "STORE_B": {
            "total": 5,
            "followedUp": 3,
            "contribution": "60.0",
            "dueToday": 1,
            "overdue": 1,
            "progress": {"TERHUBUNG": 2, "PERKENALAN": 3},
            "feedback": {"Closed": 2, "No Response": 3}
        }
    },
    "picStats": {
        "Agent Smith": {
            "total": 6,
            "followedUp": 5,
            "contribution": "83.3",
            "dueToday": 2,
            "overdue": 0,
            "progress": {"PERKENALAN": 4, "PENDEKATAN": 2},
            "feedback": {"Interested": 5, "No Response": 1}
        },
        "Agent Doe": {
            "total": 4,
            "followedUp": 2,
            "contribution": "50.0",
            "dueToday": 0,
            "overdue": 1,
            "progress": {"TERHUBUNG": 2, "PERKENALAN": 2},
            "feedback": {"Closed": 2, "No Response": 2}
        }
    }
}

def verify_dashboard_stats():
    backend_path = os.path.abspath("backend/Index.html")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Inject mock via init script
        mock_script = f"""
            window.statsData = '{json.dumps(MOCK_STATS)}';

            const runner = {{
                callback: null,
                withSuccessHandler: function(cb) {{
                    const newRunner = Object.assign({{}}, this);
                    newRunner.callback = cb;
                    return newRunner;
                }},
                withFailureHandler: function(cb) {{
                    return this;
                }},
                getDashboardStats: function() {{
                    const stats = JSON.parse(window.statsData);
                    setTimeout(() => {{
                        if (this.callback) this.callback(stats);
                    }}, 100);
                }},
                getAllClients: function() {{
                    setTimeout(() => {{
                        if (this.callback) this.callback([]);
                    }}, 100);
                }},
                getClientsNeedingFollowup: function() {{
                     setTimeout(() => {{ if (this.callback) this.callback([]); }}, 100);
                }},
                recordFollowup: function() {{}},
                exportClientData: function() {{}},
                getClientFollowupHistory: function() {{}}
            }};

            window.google = {{ script: {{ run: runner }} }};
        """
        page.add_init_script(mock_script)

        page.on("console", lambda msg: print(f"CONSOLE: {msg.text}"))

        # Load the page
        page.goto(f"file://{backend_path}")

        # Wait for stats to load
        try:
            page.wait_for_selector("#breakdownContent table", timeout=5000)
        except Exception:
            print("TIMEOUT! Dumping page content (body):")
            print(page.locator("#dashboardContent").inner_html())
            raise

        # 1. Verify Default View (Store)
        print("Verifying Store View...")
        store_btn_class = page.locator("#btnStoreStats").get_attribute("class")
        if "btn-success" not in store_btn_class:
            print("FAIL: Store button should be active by default")
            sys.exit(1)

        # Check rows in Overview table
        rows = page.locator("#breakdownContent table").nth(0).locator("tbody tr")
        if rows.count() != 2:
            print(f"FAIL: Expected 2 rows in Store Overview, got {rows.count()}")
            sys.exit(1)

        row1_text = rows.nth(0).inner_text()
        if "STORE_A" not in row1_text:
            print(f"FAIL: Expected STORE_A in first row, got {row1_text}")
            sys.exit(1)

        # Check Feedback table (3rd table)
        feed_rows = page.locator("#breakdownContent table").nth(2).locator("tbody tr")
        row1_feed = feed_rows.nth(0).inner_text()
        print(f"Feedback Row 1: {row1_feed}")

        page.screenshot(path="verification/dashboard_store.png")

        # 2. Switch to PIC
        print("Switching to PIC View...")
        page.click("#btnPicStats")

        # Verify PIC button active
        pic_btn_class = page.locator("#btnPicStats").get_attribute("class")
        if "btn-success" not in pic_btn_class:
            print("FAIL: PIC button should be active")
            sys.exit(1)

        # Check rows in Overview table
        rows = page.locator("#breakdownContent table").nth(0).locator("tbody tr")
        if rows.count() != 2: # Agent Smith, Agent Doe
             print(f"FAIL: Expected 2 rows in PIC Overview, got {rows.count()}")
             sys.exit(1)

        row1_text = rows.nth(0).inner_text()
        if "Agent" not in row1_text:
             print(f"FAIL: Expected Agent in first row, got {row1_text}")
             sys.exit(1)

        page.screenshot(path="verification/dashboard_pic.png")

        print("SUCCESS: Dashboard breakdown verified.")
        browser.close()

if __name__ == "__main__":
    verify_dashboard_stats()
