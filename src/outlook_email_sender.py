"""Outlook Web email sending via Playwright.

This module is the sending half of outwright. A long-running process
signs in to Outlook Web (see outlook_authentication) and then watches
a notification ("trigger") file. Each notification names a recipient
list, a subject, and a file holding the message body; the message is
composed and sent through the Outlook Web UI with Playwright, and the
trigger file is then removed. Malformed notification files and failed
sends are dead-lettered (renamed to ``<trigger>.bad`` /
``<trigger>.failed``) instead of being retried forever.

Key components:
    - EmailDetails: Dataclass describing one outgoing message.
    - email_send: Compose and send one message in a signed-in context.
    - emailRequests_listen: Poll the notification file and send.
    - sendOptions_parse: Parse the command-line options.
    - sender_run: Asynchronous entry point (authenticate, then listen).
    - outwright_main: Console-script entry point (see setup.py).

Typical usage:
    outwright --email me@example.com --password secret \\
        --notification /tmp/notification.txt

Notification file format (one value per line):
    line 1: comma-separated recipient addresses
    line 2: subject
    line 3: path to a file containing the message body

Dependencies:
    - playwright: Browser automation (Chromium).
    - loguru: Structured logging to stderr.
    - outlook_authentication: Browser setup and sign-in.
"""

from __future__ import annotations

import asyncio
import sys
from argparse import ArgumentDefaultsHelpFormatter, ArgumentParser, Namespace
from dataclasses import dataclass
from pathlib import Path

from loguru import logger

LOG = logger.debug
from playwright.async_api import BrowserContext, Locator, Page
from playwright.async_api import Error as PlaywrightError
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from outlook_authentication import (
    OutlookConfig,
    browser_setup,
    outlook_authenticate,
    resource_close,
)

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

notificationPollSeconds: float = 5.0
notificationLinesExpected: int = 3
peoplePickerSettleMs: int = 500
sendConfirmTimeoutMs: int = 30_000


@dataclass
class EmailDetails:
    """One outgoing email.

    Attributes:
        recipient: Comma-separated recipient addresses (To line).
        subject: Message subject.
        body_file: Path to a file containing the message body text.
        cc: Optional comma-separated Cc addresses.
    """

    recipient: str
    subject: str
    body_file: str
    cc: str = ""


def sendOptions_parse() -> Namespace:
    """Parse command-line options for the sender CLI.

    Returns:
        The parsed argument namespace (email, password, username,
        notification).
    """
    parser = ArgumentParser(
        description="An Outlook/Playwright email handler",
        formatter_class=ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--email", default="name@example.com", type=str, help="user email"
    )
    parser.add_argument(
        "--password", default="password", type=str, help="user password"
    )
    parser.add_argument(
        "--username", default="", type=str, help="internal username"
    )
    parser.add_argument(
        "--notification",
        default="/tmp/notification.txt",
        type=str,
        help="notification file",
    )
    options: Namespace = parser.parse_args()
    return options


async def outlookPage_get(context: BrowserContext) -> Page:
    """Select the page to drive: the Outlook page if one is open.

    Args:
        context: Browser context holding the signed-in session.

    Returns:
        The first page whose URL contains "outlook"; otherwise the
        most recently opened page; otherwise a fresh page.
    """
    for candidatePage in context.pages:
        if "outlook" in candidatePage.url:
            return candidatePage
    if context.pages:
        return context.pages[-1]
    return await context.new_page()


async def recipients_fill(
    page: Page, fieldSelector: str, recipientsCsv: str
) -> None:
    """Fill a recipient field, resolving each address via the picker.

    Each address is typed into the field and confirmed with Enter so
    Outlook's people picker resolves it into a recipient chip; commas
    are never typed literally into the field.

    Args:
        page: Compose page.
        fieldSelector: CSS selector of the recipient input field.
        recipientsCsv: Comma-separated recipient addresses.
    """
    for recipientAddress in recipientsCsv.split(","):
        address: str = recipientAddress.strip()
        if not address:
            continue
        await page.fill(fieldSelector, address)
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(peoplePickerSettleMs)


async def ccRecipients_fill(page: Page, ccCsv: str) -> None:
    """Fill the Cc field of an open compose form, best effort.

    The Cc input is hidden until the "Cc" toggle is clicked. A failure
    here is logged as a warning and does not fail the send itself.

    Args:
        page: Compose page.
        ccCsv: Comma-separated Cc addresses.
    """
    try:
        await page.get_by_role("button", name="Cc").click()
        await recipients_fill(page, '[aria-label="Cc"]', ccCsv)
    except PlaywrightError as error:
        LOG("Could not fill Cc recipients ({}): {}", ccCsv, error)


async def sendConfirmation_wait(
    page: Page, timeoutMs: int = sendConfirmTimeoutMs
) -> bool:
    """Wait for evidence that the message was sent.

    Success is either the compose form detaching from the page or a
    "Message sent" toast appearing, whichever happens first. The wait
    is bounded by ``timeoutMs``.

    Args:
        page: Compose page on which Send was clicked.
        timeoutMs: Maximum time to wait, in milliseconds.

    Returns:
        True if a confirmation signal was observed, False on timeout.
    """

    async def toast_wait() -> bool:
        """Wait for the 'Message sent' toast."""
        try:
            await page.wait_for_selector(
                'text="Message sent"', timeout=timeoutMs
            )
            return True
        except PlaywrightTimeoutError:
            return False

    async def compose_wait() -> bool:
        """Wait for the compose form to detach."""
        try:
            await page.wait_for_selector(
                '[placeholder="Add a subject"]',
                state="detached",
                timeout=timeoutMs,
            )
            return True
        except PlaywrightTimeoutError:
            return False

    tasks: list[asyncio.Task[bool]] = [
        asyncio.create_task(toast_wait()),
        asyncio.create_task(compose_wait()),
    ]
    done, pending = await asyncio.wait(
        tasks, return_when=asyncio.FIRST_COMPLETED
    )
    for task in pending:
        task.cancel()
    if pending:
        await asyncio.gather(*pending, return_exceptions=True)
    return any(task.result() for task in done)


async def email_send(
    context: BrowserContext, options: Namespace, emailDetails: EmailDetails
) -> bool:
    """Compose and send one email in a signed-in browser context.

    The send is only reported successful once a confirmation signal
    (compose form gone, or a "Message sent" toast) has been observed;
    this verification runs in every sign-in mode.

    Args:
        context: Browser context holding the signed-in session.
        options: Sender CLI options (used for the sign-in mode flag).
        emailDetails: Recipients, subject, and body file to send.

    Returns:
        True if the message was sent and confirmed, False otherwise.
    """
    try:
        page: Page = await outlookPage_get(context)

        LOG("Clicking 'New mail'")
        await page.click('[aria-label="New mail"]')
        await page.wait_for_selector('[aria-label="To"]')

        LOG("Filling email details")
        LOG("Internal sign-in mode: {}", bool(options.username))
        await recipients_fill(page, '[aria-label="To"]', emailDetails.recipient)
        if emailDetails.cc.strip():
            await ccRecipients_fill(page, emailDetails.cc)
        await page.fill(
            '[placeholder="Add a subject"]', emailDetails.subject.strip()
        )

        LOG("Adding email body")
        bodyText: str = await asyncio.to_thread(
            Path(emailDetails.body_file).read_text, encoding="utf-8"
        )
        bodyEditor: Locator = page.locator(
            'div[aria-label="Message body"][contenteditable="true"]'
        )
        await bodyEditor.click()
        await bodyEditor.fill(bodyText.strip())

        LOG("Sending email")
        await page.click('[aria-label="Send"]')

        LOG("Waiting for email sent confirmation")
        if not await sendConfirmation_wait(page):
            LOG(
                "No send confirmation observed; "
                "the email may not have been sent."
            )
            return False

        LOG("Email sent to {}.", emailDetails.recipient)
        return True
    except (PlaywrightError, OSError) as error:
        LOG("Error sending email: {}", error)
        return False


async def emailRequests_listen(
    context: BrowserContext, options: Namespace, triggerFile: Path
) -> None:
    """Poll the notification file and send whatever it describes.

    A well-formed notification file is removed after a successful
    send. A malformed file (fewer than three lines) is renamed to
    ``<trigger>.bad`` and a failed send to ``<trigger>.failed`` so a
    poison notification is never retried in a loop.

    Args:
        context: Browser context holding the signed-in session.
        options: Sender CLI options.
        triggerFile: Path of the notification file to watch.
    """
    LOG("Waiting for email notifications...")
    while True:
        try:
            if triggerFile.is_file():
                LOG("Notification file found: {}", triggerFile)
                notificationText: str = await asyncio.to_thread(
                    triggerFile.read_text, encoding="utf-8"
                )
                notificationLines: list[str] = (
                    notificationText.splitlines()
                )
                if len(notificationLines) < notificationLinesExpected:
                    LOG(
                        "Malformed notification file (expected {} lines, "
                        "got {}): {}. Moving it aside.",
                        notificationLinesExpected,
                        len(notificationLines),
                        triggerFile,
                    )
                    await asyncio.to_thread(
                        triggerFile.replace, Path(f"{triggerFile}.bad")
                    )
                else:
                    emailDetails: EmailDetails = EmailDetails(
                        recipient=notificationLines[0].strip(),
                        subject=notificationLines[1].strip(),
                        body_file=notificationLines[2].strip(),
                    )
                    emailSent: bool = await email_send(
                        context, options, emailDetails
                    )
                    if emailSent:
                        await asyncio.to_thread(triggerFile.unlink)
                        LOG(
                            "Notification file processed and removed: {}",
                            triggerFile,
                        )
                    else:
                        LOG(
                            "Failed to send email. Moving notification "
                            "file aside: {}",
                            triggerFile,
                        )
                        await asyncio.to_thread(
                            triggerFile.replace,
                            Path(f"{triggerFile}.failed"),
                        )
        except (OSError, PlaywrightError, UnicodeDecodeError) as error:
            LOG("Error in emailRequests_listen: {}", error)

        await asyncio.sleep(notificationPollSeconds)


async def sender_run() -> None:
    """Run the sender: authenticate, then listen for notifications.

    Browser resources are closed individually on exit, whatever stage
    the run reached.
    """
    options: Namespace = sendOptions_parse()
    config: OutlookConfig = OutlookConfig(
        email=options.email,
        password=options.password,
        username=options.username,
    )

    playwright, browser, context = await browser_setup()

    if playwright is None or browser is None or context is None:
        LOG("Failed to set up browser. Exiting.")
        return

    try:
        LOG("Starting authentication process...")
        authenticated: bool = await outlook_authenticate(context, config)
        if authenticated:
            LOG(
                "Authentication successful. "
                "The script will now run indefinitely."
            )
            LOG("Press Ctrl+C to stop the script at any time.")
            await emailRequests_listen(
                context, options, Path(options.notification)
            )
        else:
            LOG("Failed to authenticate. Exiting.")
    except KeyboardInterrupt:
        LOG("\nScript terminated by user.")
    except (PlaywrightError, OSError) as error:
        LOG("An error occurred in sender_run: {}", error)
    finally:
        LOG("Cleaning up resources...")
        await resource_close(context.close, "browser context")
        await resource_close(browser.close, "browser")
        await resource_close(playwright.stop, "playwright driver")


def outwright_main() -> None:
    """Console-script entry point for the outwright sender."""
    asyncio.run(sender_run())


if __name__ == "__main__":
    outwright_main()
