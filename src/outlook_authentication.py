"""Outlook Web authentication via Playwright.

This module launches a Chromium browser through Playwright and signs a
user into Outlook Web (https://outlook.office.com). It supports two
sign-in modes: a direct credential fill for plain Microsoft accounts,
and an "internal" mode (used behind institutional single sign-on, e.g.
a hospital IdP) where the human completes the sign-in interactively in
the browser window while this module simply waits for the Outlook
mailbox UI to appear.

Key components:
    - OutlookConfig: Dataclass holding the sign-in parameters.
    - browser_setup: Launch Chromium and create a browser context.
    - outlook_authenticate: Drive the Outlook Web sign-in flow.
    - resource_close: Best-effort close of a partially started resource.
    - authOptions_parse: Parse the command-line options.
    - authentication_run: Asynchronous CLI entry point.

Typical usage:
    python outlook_authentication.py --email me@example.com \\
        --password secret

Dependencies:
    - playwright: Browser automation (Chromium).
    - loguru: Structured logging to stderr.
"""

from __future__ import annotations

import asyncio
import sys
from argparse import ArgumentDefaultsHelpFormatter, ArgumentParser, Namespace
from collections.abc import Awaitable, Callable
from dataclasses import dataclass

from loguru import logger

LOG = logger.debug
from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    async_playwright,
)
from playwright.async_api import Error as PlaywrightError
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

loggerFormat: str = (
    "<green>{time:YYYY-MM-DD HH:mm:ss}</green> │ "
    "<level>{level: <5}</level> │ "
    "<yellow>{name: >28}</yellow>::"
    "<cyan>{function: <30}</cyan> @"
    "<cyan>{line: <4}</cyan> ║ "
    "<level>{message}</level>"
)
logger.remove()
logger.add(sys.stderr, format=loggerFormat)

defaultTimeoutMs: int = 30_000
authWaitTimeoutMs: int = 60_000


@dataclass
class OutlookConfig:
    """Configuration for an Outlook Web sign-in.

    Attributes:
        email: Account email address used to sign in.
        password: Account password. Only used in the direct sign-in
            mode; in internal mode the human signs in interactively
            and this value is never typed anywhere.
        username: Internal (institutional) username. When non-empty,
            authentication runs in internal mode: the human completes
            the sign-in by hand in the browser window.
        outlook_url: Outlook Web entry point URL.
    """

    email: str
    password: str
    username: str
    outlook_url: str = "https://outlook.office.com"


async def resource_close(
    closeAction: Callable[[], Awaitable[None]], resourceName: str
) -> None:
    """Close a browser resource, logging (not raising) any failure.

    Args:
        closeAction: Zero-argument coroutine function that closes the
            resource, e.g. ``context.close``.
        resourceName: Human-readable resource name for log messages.
    """
    try:
        await closeAction()
    except PlaywrightError as error:
        LOG("Failed to close {}: {}", resourceName, error)


async def browser_setup() -> (
    tuple[Playwright | None, Browser | None, BrowserContext | None]
):
    """Launch Chromium and create a browser context.

    Returns:
        A ``(playwright, browser, context)`` tuple. On failure the
        partially started resources are closed individually and the
        tuple is ``(None, None, None)``.
    """
    playwright: Playwright | None = None
    browser: Browser | None = None
    context: BrowserContext | None = None
    try:
        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(headless=False)
        context = await browser.new_context(no_viewport=True)
        context.set_default_timeout(defaultTimeoutMs)
        LOG("Browser context created.")
        return playwright, browser, context
    except PlaywrightError as error:
        LOG("Error setting up browser: {}", error)
        if context is not None:
            await resource_close(context.close, "browser context")
        if browser is not None:
            await resource_close(browser.close, "browser")
        if playwright is not None:
            await resource_close(playwright.stop, "playwright driver")
        return None, None, None


async def outlook_authenticate(
    context: BrowserContext, config: OutlookConfig
) -> bool:
    """Sign in to Outlook Web in a fresh page of the given context.

    In internal mode (``config.username`` non-empty) the human is
    expected to complete the institutional sign-in interactively; this
    function then only waits for the mailbox UI. Otherwise the email
    and password are filled programmatically and the optional
    "Is this your device?" / "Stay signed in?" prompts are accepted.

    Args:
        context: Browser context to open the sign-in page in.
        config: Sign-in parameters.

    Returns:
        True if the mailbox UI (the "New mail" button) became visible,
        False on any failure.
    """
    try:
        page: Page = await context.new_page()

        LOG("Navigating to {}", config.outlook_url)
        await page.goto(config.outlook_url)

        LOG("Filling email")
        await page.fill('input[type="email"]', config.email)
        await page.click('input[type="submit"]')

        if config.username:
            LOG(
                "Internal authentication -- "
                "please handle the authentication manually"
            )
        else:
            await page.wait_for_selector(
                "#userNameInput", timeout=authWaitTimeoutMs
            )
            await page.fill("#userNameInput", config.email)
            await page.fill("#passwordInput", config.password)
            await page.click("#submitButton")

            LOG("Checking for 'Is this your device?' prompt")
            try:
                await page.wait_for_selector(
                    "#trust-browser-button", timeout=authWaitTimeoutMs
                )
                await page.click("#trust-browser-button")
                LOG("Clicked 'Yes, this is my device'")
            except PlaywrightTimeoutError:
                LOG(
                    "No 'Is this your device?' prompt detected. "
                    "Continuing..."
                )

            LOG("Checking for 'Stay signed in?' prompt")
            try:
                await page.wait_for_selector(
                    "#idSIButton9", timeout=authWaitTimeoutMs
                )
                await page.click("#idSIButton9")
                LOG("Clicked 'Stay signed in'")
            except PlaywrightTimeoutError:
                LOG(
                    "No 'Stay signed in?' prompt detected. Continuing..."
                )

        LOG("Waiting for the final page to load")
        await page.wait_for_selector(
            '[aria-label="New mail"]', timeout=authWaitTimeoutMs
        )
        LOG("Authenticated successfully.")
        return True
    except PlaywrightError as error:
        LOG("Error during authentication: {}", error)
        return False


def authOptions_parse() -> Namespace:
    """Parse command-line options for the authentication CLI.

    Returns:
        The parsed argument namespace (email, password, username).
    """
    parser = ArgumentParser(
        description="An Outlook/Playwright authenticator",
        formatter_class=ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--email", default="", type=str, help="user email")
    parser.add_argument(
        "--password", default="password", type=str, help="user password"
    )
    parser.add_argument(
        "--username", default="", type=str, help="internal username"
    )

    options: Namespace = parser.parse_args()
    return options


async def authentication_run() -> None:
    """Run the authentication CLI: set up a browser and sign in.

    Whatever browser resources were started are closed individually on
    exit, even if setup only partially succeeded.
    """
    options: Namespace = authOptions_parse()
    config: OutlookConfig = OutlookConfig(
        email=options.email,
        password=options.password,
        username=options.username,
    )
    playwright, browser, context = await browser_setup()
    try:
        if context is None:
            LOG("Browser context was not defined. Critical error")
            return
        authenticated: bool = await outlook_authenticate(context, config)
        if authenticated:
            LOG(
                "Authentication was successful. "
                "You can proceed with sending emails."
            )
        else:
            LOG(
                "Authentication failed. "
                "Please check your credentials and try again."
            )
    finally:
        if context is not None:
            await resource_close(context.close, "browser context")
        if browser is not None:
            await resource_close(browser.close, "browser")
        if playwright is not None:
            await resource_close(playwright.stop, "playwright driver")


if __name__ == "__main__":
    asyncio.run(authentication_run())
