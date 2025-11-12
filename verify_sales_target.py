
import asyncio
from playwright.async_api import async_playwright
import json
import os

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()

        # Get the absolute path to the HTML file
        html_file_path = os.path.abspath('Index.html')

        # Go to the local HTML file
        await page.goto(f'file://{html_file_path}')

        # Read the mock data from the JSON file
        with open('mock_data.json', 'r') as f:
            mock_data = json.load(f)

        # Inject the mock data into the page and call the updateDashboard function
        await page.evaluate(f'window.updateDashboard({json.dumps(mock_data)})')

        # Wait for the chart animation to complete
        await page.wait_for_timeout(2000)

        # Take a screenshot of the overview page
        await page.screenshot(path='overview_with_targets.png')

        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
