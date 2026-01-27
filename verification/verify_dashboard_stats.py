from playwright.sync_api import sync_playwright
import os
import json

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()
    page.set_viewport_size({"width": 1280, "height": 800})

    # Listen for console logs
    page.on("console", lambda msg: print(f"Console: {msg.text}"))
    page.on("pageerror", lambda err: print(f"Page Error: {err}"))

    # Mock Data
    mock_dashboard = {
        "totalClients": 10,
        "dueFollowups": 2,
        "overdueClients": 1,
        "avgFollowUps": 1.5,
        "followUpPercentage": 60.0,
        "progressCounts": {
            "PERKENALAN": 5,
            "PENDEKATAN": 3,
            "TERHUBUNG": 2
        },
        "feedbackCounts": {
            "SAY THANKYOU": 4,
            "NO RESPON": 2,
            "BERSEDIA DIFOLLOW UP LEBIH LANJUT": 4
        }
    }

    dashboard_json = json.dumps(mock_dashboard)

    # Improved Mock to handle chained callbacks correctly
    page.add_init_script(f"""
        const runner = {{
            withSuccessHandler: function(cb) {{
                const next = Object.create(this);
                next._success = cb;
                return next;
            }},
            withFailureHandler: function(cb) {{
                const next = Object.create(this);
                next._failure = cb;
                return next;
            }},
            getAllClients: function() {{
                const cb = this._success;
                setTimeout(() => cb && cb([]), 100);
            }},
            getDashboardStats: function() {{
                const cb = this._success;
                const data = {dashboard_json};
                console.log("Mock returning dashboard data");
                setTimeout(() => cb && cb(data), 100);
            }},
            getClientsNeedingFollowup: function() {{
                const cb = this._success;
                setTimeout(() => cb && cb([]), 100);
            }},
            getClientFollowupHistory: function(row) {{
                 const cb = this._success;
                 setTimeout(() => cb && cb([]), 100);
            }}
        }};

        window.google = {{
            script: {{
                run: runner
            }}
        }};
    """)

    cwd = os.getcwd()
    file_url = f"file://{cwd}/frontend/index.html"
    print(f"Loading {file_url}")
    page.goto(file_url)

    print("Waiting for dashboard to render...")
    try:
        page.wait_for_selector("text=Follow-up Contribution", timeout=5000)
        page.wait_for_selector("text=Progress Breakdown", timeout=5000)
        page.wait_for_selector("text=Feedback Breakdown", timeout=5000)

        # Check for specific values
        page.wait_for_selector("text=60%", timeout=5000)
        page.wait_for_selector("text=PERKENALAN", timeout=5000)
        page.wait_for_selector("text=SAY THANKYOU", timeout=5000)

        print("Dashboard elements verified.")
        page.screenshot(path="verification/verification_dashboard_stats.png")
        print("Screenshot saved.")
    except Exception as e:
        print(f"Verification failed: {e}")
        page.screenshot(path="verification/failed_verification.png")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
