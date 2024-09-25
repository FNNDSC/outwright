# outlook_email_sender.py
import os
import asyncio
from dataclasses import dataclass
from pathlib import Path
from playwright.async_api import BrowserContext, Page
from typing import Optional
from argparse import ArgumentParser, ArgumentDefaultsHelpFormatter, Namespace
from outlook_authentication import setup_browser, authenticate_outlook, OutlookConfig
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
class EmailDetails:
    recipient: str
    subject: str
    body_file: str


def send_argparse() -> Namespace:
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
        "--notification",
        default="/tmp/notification.txt",
        type=str,
        help="notification file",
    )
    options: Namespace = parser.parse_args()
    return options


async def send_email(context: BrowserContext, email_details: EmailDetails) -> bool:
    try:
        page = context.pages[0]  # Use the existing page

        LOG("Clicking 'New mail'")
        await page.click('[aria-label="New mail"]')
        await page.wait_for_selector('[aria-label="To"]')

        LOG("Filling email details")
        recipients = email_details.recipient.split(",")
        for i, recipient in enumerate(recipients):
            await page.fill('[aria-label="To"]', recipient.strip())
            if i < len(recipients) - 1:  # If it's not the last recipient
                await page.keyboard.type(",")  # Add an explicit comma
                await page.keyboard.press("Space")  # Add a space after the comma

        await page.fill('[placeholder="Add a subject"]', email_details.subject.strip())
        await page.keyboard.press("Tab")  # Move to the body

        LOG("Adding email body")
        with open(email_details.body_file, "r") as file:
            body: str = file.read().strip()
        await page.keyboard.type(body)

        LOG("Sending email")
        await page.keyboard.press("Control+Enter")

        LOG("Waiting for email sent confirmation")
        await page.wait_for_selector("#EmptyState_MainMessage", timeout=30000)

        LOG(f"Email sent to {email_details.recipient}.")
        return True
    except Exception as e:
        LOG(f"Error sending email: {e}")
        return False


async def listen_for_email_requests(context: BrowserContext, triggerFile: Path) -> None:
    LOG("Waiting for email notifications...")
    while True:
        try:
            if triggerFile.is_file():
                LOG(f"Notification file found: {triggerFile}")
                with open(triggerFile, "r") as file:
                    lines: list[str] = file.readlines()
                    recipient: str = lines[0].strip()
                    subject: str = lines[1].strip()
                    body_file: str = lines[2].strip()

                email_details: EmailDetails = EmailDetails(
                    recipient=recipient, subject=subject, body_file=body_file
                )

                email_sent: bool = await send_email(context, email_details)
                if email_sent:
                    os.remove(triggerFile)
                    LOG(f"Notification file processed and removed: {triggerFile}")
                else:
                    LOG(
                        f"Failed to send email. Notification file not removed: {triggerFile}"
                    )

        except Exception as e:
            LOG(f"Error in listen_for_email_requests: {e}")

        await asyncio.sleep(5)  # Check for notifications every 5 seconds


async def async_main() -> None:
    options: Namespace = send_argparse()
    config: OutlookConfig = OutlookConfig(
        email=options.email,
        password=options.password,
    )

    playwright, browser, context = await setup_browser()

    if playwright is None or browser is None or context is None:
        LOG("Failed to set up browser. Exiting.")
        return

    try:
        LOG("Starting authentication process...")
        authenticated: bool = await authenticate_outlook(context, config)
        if authenticated:
            LOG("Authentication successful. The script will now run indefinitely.")
            LOG("Press Ctrl+C to stop the script at any time.")
            await listen_for_email_requests(context, Path(options.notification))
        else:
            LOG("Failed to authenticate. Exiting.")
    except KeyboardInterrupt:
        LOG("\nScript terminated by user.")
    except Exception as e:
        LOG(f"An error occurred in async_main: {e}")
    finally:
        LOG("Cleaning up resources...")
        if context:
            await context.close()
        if browser:
            await browser.close()
        if playwright:
            await playwright.stop()


def main() -> None:
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
