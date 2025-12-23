import os
import sys

from dotenv import load_dotenv
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


def login(page: Page, email: str, password: str) -> None:
    """Sign in to Outlook contacts."""
    page.goto("https://outlook.office.com/mail/0/?deeplink=mail%2F0%2F", wait_until="domcontentloaded")

    try:
        sign_in_link = page.get_by_role("link", name="Sign in")
        if sign_in_link.is_visible(timeout=5000):
            sign_in_link.click()
    except PlaywrightTimeoutError:
        pass

    page.get_by_label("Email, phone, or Skype").fill(email)
    page.get_by_role("button", name="Next").click()

    page.get_by_label("Password").fill(password)
    page.get_by_role("button", name="Sign in").click()

    try:
        # Dismiss the "Stay signed in?" prompt if it appears.
        page.get_by_role("button", name="No").click(timeout=4000)
    except PlaywrightTimeoutError:
        pass

    # page.wait_for_url("https://outlook.live.com/people/*", wait_until="domcontentloaded")


def delete_contact(page: Page) -> None:
    """Delete a contact by name from the contacts list."""
    # Navigate to People/Contacts in case we are redirected elsewhere.
    # page.goto("https://outlook.office.com/mail/0/?deeplink=mail%2F0%2F", wait_until="aria-label="People")")

    # # Search the contact.
    # search_box = page.get_by_role("combobox", name="Search")
    # search_box.click()
    # search_box.fill(contact_name)
    # search_box.press("Enter")

    # Open the first matching contact.
    contact_row = page.locator("xpath=//button[@aria-label='People']").first
    contact_row.click()

    # Open the command bar and delete the contact.
    contact_list = page.locator("xpath=//span[text()='Your contact lists']").first
    contact_list.click()

    # Try a few times in case the delete/confirm buttons appear at different times.
    count = 0
    while True:

        delete_button = page.locator("xpath=//span[text()='Delete']")
        confirm_button = page.locator("xpath=//button[text()='Delete']")


        try:
            delete_button.wait_for(state="visible", timeout=5000)
            delete_button.click(timeout=5000)
        except PlaywrightTimeoutError:
            # Delete button did not show in time; try confirm anyway.
            continue

        try:
            confirm_button.wait_for(state="visible", timeout=5000)
            confirm_button.click(timeout=5000)
        except PlaywrightTimeoutError:
            # If confirm is missing, wait a bit and retry.
            page.wait_for_timeout(1000)
        count += 1
        if count >= 5:
            page.reload()
            contact_list = page.locator("xpath=//span[text()='Your contact lists']").first
            contact_list.click()
            count = 0



def run(headless: bool = True, browser_name: str = "chromium") -> None:
    load_dotenv()

    # email = os.getenv("OUTLOOK_EMAIL")
    # password = os.getenv("OUTLOOK_PASSWORD")
    email = "automation_01@nakivo04.onmicrosoft.com"
    password = "Performance@123"
    if not email or not password:
        raise RuntimeError("Please set OUTLOOK_EMAIL and OUTLOOK_PASSWORD in your environment or .env file.")

    with sync_playwright() as p:
        browser_name = (os.getenv("PLAYWRIGHT_BROWSER") or browser_name or "chromium").strip().lower()
        if browser_name not in {"chromium", "firefox", "webkit"}:
            raise ValueError("browser_name must be one of: chromium, firefox, webkit")

        browser_type = getattr(p, browser_name)
        browser = browser_type.launch(headless=headless)
        while True:
            context = None
            try:
                context = browser.new_context()
                page = context.new_page()

                login(page, email, password)
                delete_contact(page)
                break  # Success
            except Exception:
                # Reset session/cache and retry from a fresh login.
                if context:
                    context.close()
                continue
            finally:
                if context:
                    context.close()
        browser.close()


if __name__ == "__main__":
    browser_name = "chromium"
    if "--browser" in sys.argv:
        idx = sys.argv.index("--browser")
        if idx + 1 < len(sys.argv):
            browser_name = sys.argv[idx + 1]

    run(headless=False, browser_name=browser_name)
