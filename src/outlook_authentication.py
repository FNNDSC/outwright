# outlook_authentication.py
import asyncio
from dataclasses import dataclass
from playwright.async_api import (
    async_playwright,
    BrowserContext,
    Browser,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
)

from typing import Optional
from argparse import ArgumentParser, ArgumentDefaultsHelpFormatter, Namespace
import sys
from loguru import logger

LOG = logger.debug
logger_format = (
    "<green>{time:YYYY-MM-DD HH:mm:ss}</green> │ "
    "<level>{level: <5}</level> │ "
    "<yellow>{name: >28}</yellow>::"
    "<cyan>{function: <30}</cyan> @"
    "<cyan>{line: <4}</cyan> ║ "
    "<level>{message}</level>"
)
logger.remove()
logger.add(sys.stderr, format=logger_format)


@dataclass
class OutlookConfig:
    email: str
    password: str
    username: str
    outlook_url: str = "https://outlook.office.com"


async def setup_browser() -> (
    tuple[Optional[Playwright], Optional[Browser], Optional[BrowserContext]]
):
    try:
        playwright: Playwright = await async_playwright().start()
        browser: Browser = await playwright.chromium.launch(headless=False)
        context: BrowserContext = await browser.new_context(no_viewport=True)

        if context is not None:
            context.set_default_timeout(2147483647)  # Maximum 32-bit integer value
            LOG("Browser context created with extended timeout.")
        else:
            LOG("Failed to create browser context.")

        return playwright, browser, context
    except Exception as e:
        LOG(f"Error setting up browser: {e}")
        return None, None, None


async def authenticate_outlook(context: BrowserContext, config: OutlookConfig) -> bool:
    try:
        page: Page = await context.new_page()

        LOG(f"Navigating to {config.outlook_url}")
        await page.goto(config.outlook_url)

        LOG("Filling email")
        await page.fill('input[type="email"]', config.email)
        await page.click('input[type="submit"]')

        if config.username:
            LOG("Internal authentication -- please handle the authentication manually")
        else:
            LOG("Using external authentication flow")
            LOG("Waiting for password input")
            await page.wait_for_selector('input[type="password"]', timeout=60000)
            await page.fill('input[type="password"]', config.password)
            await page.click('input[type="submit"]')

            LOG("Checking for 'Stay signed in?' prompt")
            try:
                await page.wait_for_selector("#idSIButton9", timeout=60000)
                await page.click("#idSIButton9")
                LOG("Clicked 'Stay signed in'")
            except PlaywrightTimeoutError:
                LOG("No 'Stay signed in?' prompt detected. Continuing...")

        LOG("Waiting for the final page to load")
        await page.wait_for_selector('[aria-label="New mail"]', timeout=60000)
        LOG("Authenticated successfully.")
        return True
    except Exception as e:
        LOG(f"Error during authentication: {e}")
        return False


def auth_argparse() -> Namespace:
    parser = ArgumentParser(
        description="An Outlook/Playwright authenticator",
        formatter_class=ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--email", default="", type=str, help="user email")
    parser.add_argument(
        "--password", default="password", type=str, help="user password"
    )
    parser.add_argument("--username", default="", type=str, help="internal username")

    options: Namespace = parser.parse_args()
    return options


async def authenticate() -> None:
    options: Namespace = auth_argparse()
    config: OutlookConfig = OutlookConfig(
        email=options.email, password=options.password, username=options.username
    )
    playwright, browser, context = await setup_browser()
    try:
        if not context:
            LOG("Browser context was not defined. Critical error")
            return
        authenticated: bool = await authenticate_outlook(context, config)
        if authenticated:
            LOG("Authentication was successful. You can proceed with sending emails.")
        else:
            LOG("Authentication failed. Please check your credentials and try again.")
    finally:
        if context and browser and playwright:
            await context.close()
            await browser.close()
            await playwright.stop()


if __name__ == "__main__":
    asyncio.run(authenticate())
