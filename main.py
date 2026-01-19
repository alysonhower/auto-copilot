import asyncio
import random
from pathlib import Path

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions

# Persistent Chrome profile directory (stores login sessions, cookies, etc.)
USER_DATA_DIR = Path(__file__).parent / '.chrome_profile'


async def copilot_chat_automation():
    """Type a message in Microsoft 365 Copilot chat and send it with human-like behavior."""
    # Configure Chrome to use persistent profile
    options = ChromiumOptions()
    options.add_argument(f'--user-data-dir={USER_DATA_DIR}')

    async with Chrome(options=options) as browser:
        tab = await browser.start()

        # Navigate to Microsoft 365 Copilot
        await tab.go_to('https://m365.cloud.microsoft/chat')
        
        # Human-like wait for page to load
        await asyncio.sleep(random.uniform(2.0, 4.0))

        # Find the chat input element using aria-label
        chat_input = await tab.find(
            aria_label="Copilot de Mensagens",
            timeout=10
        )

        # Click on the chat input to focus it (with human-like offset)
        await chat_input.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15)
        )
        
        # Small delay before typing (human hesitation)
        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Type message with humanized behavior (variable speed, occasional typos)
        message = "Olá!"
        await chat_input.type_text(message, humanize=True)

        # Wait a bit after typing (human reaction time before clicking send)
        await asyncio.sleep(random.uniform(0.5, 1.5))

        # Find the send button using aria-label (it appears after typing)
        send_button = await tab.find(
            aria_label="Enviar",
            timeout=5
        )

        # Click send button with human-like parameters
        await send_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.09, 0.18)
        )

        # Wait to observe the result
        await asyncio.sleep(random.uniform(2.0, 4.0))


asyncio.run(copilot_chat_automation())