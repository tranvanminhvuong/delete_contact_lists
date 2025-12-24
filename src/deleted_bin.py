import os
import sys
from typing import Optional

from dotenv import load_dotenv
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


def login(page: Page, email: str, password: str) -> None:
    """Sign in to Outlook contacts."""
    page.goto("https://outlook.office.com/mail/0/?deeplink=mail%2F0%2F", wait_until="domcontentloaded")

    try:
        sign_in_link = page.get_by_role("link", name="Sign in")
        if sign_in_link.is_visible(timeout=60000):
            sign_in_link.click()
    except PlaywrightTimeoutError:
        pass

    page.get_by_label("Email, phone, or Skype").fill(email)
    page.get_by_role("button", name="Next").click()

    page.get_by_label("Password").fill(password)
    page.get_by_role("button", name="Sign in").click()

    try:
        # Dismiss the "Stay signed in?" prompt if it appears.
        page.get_by_role("button", name="Yes").click(timeout=4000)
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
    
    contact_row.wait_for(state="visible", timeout=60000)
    contact_row.click()

    # Open the command bar and delete the contact.
    contact_list = page.locator("xpath=//span[text()='Deleted']").first
    contact_list.wait_for(state="visible", timeout=60000)
    contact_list.click()

    # Try a few times in case the delete/confirm buttons appear at different times.
    count = 0
    while True:

        delete_button = page.locator("xpath=//span[text()='Delete']")
        confirm_button = page.locator("//div[contains(@class,'Dialog')]//button//span[text()='Delete']")


        try:
            delete_button.wait_for(state="visible", timeout=60000)
            delete_button.click(timeout=60000)
        except PlaywrightTimeoutError:
            # Delete button did not show in time; try confirm anyway.
            page.reload()
            contact_list.wait_for(state="visible", timeout=60000)
            contact_list.click()
            continue

        try:
            confirm_button.wait_for(state="visible", timeout=60000)
            confirm_button.click(timeout=60000)
        except PlaywrightTimeoutError:
            # If confirm is missing, wait a bit and retry.
            page.reload()
            contact_list.wait_for(state="visible", timeout=60000)
            contact_list.click()
            continue
        count += 1
        if count >= 5:
            page.reload()
            contact_list = page.locator("xpath=//span[text()='Deleted']").first
            contact_list.wait_for(state="visible", timeout=60000)
            contact_list.click()
            count = 0



def run(contact_name: str, headless: bool = True) -> None:
    load_dotenv()

    # email = os.getenv("OUTLOOK_EMAIL")
    # password = os.getenv("OUTLOOK_PASSWORD")
    email = "automation_01@nakivo04.onmicrosoft.com"
    password = "Performance@123"
    if not email or not password:
        raise RuntimeError("Please set OUTLOOK_EMAIL and OUTLOOK_PASSWORD in your environment or .env file.")

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=headless)
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
    if len(sys.argv) < 2:
        print("Usage: python src/test_outlook_contacts.py '<CONTACT_NAME>'")
        sys.exit(1)

    contact = sys.argv[1]
    run(contact_name=contact, headless=False)
