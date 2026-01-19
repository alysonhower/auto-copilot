import asyncio
import json
import random
from pathlib import Path

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions

# Persistent cookies file (stores login sessions)
COOKIE_FILE = Path(__file__).parent / '.cookies.json'


async def load_cookies(tab):
    """Load saved cookies to restore login session."""
    if not COOKIE_FILE.exists():
        return False
    
    try:
        saved_cookies = json.loads(COOKIE_FILE.read_text(encoding='utf-8'))
        
        # Convert to simplified format (only settable fields)
        # get_cookies() returns detailed Cookie objects with read-only fields
        # set_cookies() expects CookieParam format
        simplified_cookies = []
        for cookie in saved_cookies:
            simplified = {
                'name': cookie['name'],
                'value': cookie['value'],
                'domain': cookie.get('domain'),
                'path': cookie.get('path', '/'),
                'secure': cookie.get('secure', False),
                'httpOnly': cookie.get('httpOnly', False),
            }
            # Only include expiration if present and valid
            if 'expires' in cookie and cookie['expires'] > 0:
                simplified['expires'] = cookie['expires']
            simplified_cookies.append(simplified)
        
        await tab.set_cookies(simplified_cookies)
        print(f"Loaded {len(simplified_cookies)} cookies from {COOKIE_FILE}")
        return True
    except Exception as e:
        print(f"Error loading cookies: {e}")
        return False


async def save_cookies(browser):
    """Save cookies after successful login for future use."""
    try:
        cookies = await browser.get_cookies()
        COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding='utf-8')
        print(f"Saved {len(cookies)} cookies to {COOKIE_FILE}")
    except Exception as e:
        print(f"Error saving cookies: {e}")


async def copilot_chat_automation():
    """Type a message in Microsoft 365 Copilot chat and send it with human-like behavior."""
    # Configure Chrome with preferences
    options = ChromiumOptions()
    
    # Block notifications and popups for cleaner automation
    options.block_notifications = True
    options.block_popups = True

    async with Chrome(options=options) as browser:
        tab = await browser.start()
        
        # Load existing cookies to restore login session
        await load_cookies(tab)

        # Navigate to Microsoft 365 Copilot
        await tab.go_to('https://m365.cloud.microsoft/chat')
        
        # Human-like wait for page to load
        await asyncio.sleep(random.uniform(2.0, 4.0))

        # === FILE ATTACHMENT ===
        # Click the plus menu button to open attachment options
        plus_menu_btn = await tab.find(data_testid="PlusMenuButton", timeout=10)
        await plus_menu_btn.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-5, 5),
            hold_time=random.uniform(0.08, 0.15)
        )

        # Human-like wait for menu to appear
        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Attach the file using file chooser context manager
        file_path = Path(r"C:\Users\AlysonhowerVerasViei\Downloads\markdown.md")
        
        async with tab.expect_file_chooser(files=[file_path]):
            upload_menu_item = await tab.find(text="Carregar imagens e arquivos", timeout=5)
            await upload_menu_item.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-3, 3),
                hold_time=random.uniform(0.08, 0.15)
            )

        # Wait for file attachment to process
        await asyncio.sleep(random.uniform(0.8, 1.5))

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
        message = """Analise o documento em anexo e execute as duas tarefas abaixo:

1. Produza um resumo em um único parágrafo, claro e objetivo.
2. Em seguida, identifique e destaque o objeto central do documento em uma única sentença curta, no estilo *punchline* (poucas palavras, direto ao ponto, refletindo o cerne do conteúdo).

Retorne exclusivamente no formato abaixo:

```markdown
<resumo>
[Resumo em um único parágrafo]
</resumo>

<objeto>
[Objeto central do documento]
</objeto>
"""
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

        # Wait for the LLM to generate a response
        # The copy button only appears when the model finishes generating
        # Using a long timeout since Copilot may take a while to process
        print("Aguardando resposta do Copilot...")
        
        # Wait for the copy button to appear (indicates response is complete)
        await tab.find(
            aria_label="Copiar Resposta",
            timeout=120,  # Wait up to 2 minutes for response
        )

        # Human-like delay before reading the response
        await asyncio.sleep(random.uniform(0.5, 1.5))

        # Find the response element using data-testid attribute
        # Note: use data_testid (underscore) as keyword argument per Pydoll docs
        response_element = await tab.find(
            data_testid='lastChatMessage',
            timeout=10
        )

        # Get text content using Pydoll's native element.text property
        response_text = await response_element.text

        if response_text:
            # Generate timestamped filename
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = Path(__file__).parent / 'outputs' / f'copilot_response_{timestamp}.md'
            
            # Ensure output directory exists
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Save the response to markdown file
            output_file.write_text(response_text, encoding='utf-8')
            print(f"Resposta salva em: {output_file}")
        else:
            print("Não foi possível obter o conteúdo da resposta.")

        # Save cookies for next run (persists login session)
        await save_cookies(browser)

        # Keep browser open for debugging - press Enter to close
        # NOTE: On first run, you need to log in manually. Cookies will be saved
        # and subsequent runs will restore your login session automatically.
        input("Pressione Enter para fechar o navegador...")


asyncio.run(copilot_chat_automation())