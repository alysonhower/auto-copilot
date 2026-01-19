import asyncio
import json
import random
import re
from pathlib import Path
from typing import List

from openpyxl import Workbook, load_workbook

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions

# Persistent cookies file (stores login sessions)
COOKIE_FILE = Path(__file__).parent / '.cookies.json'

# Global lock for Excel writing to prevent race conditions
EXCEL_LOCK = asyncio.Lock()

# List of files to process
FILES_TO_PROCESS = [
    Path(r"C:\Users\AlysonhowerVerasViei\Downloads\markdown.md"),
    Path(r"C:\Users\AlysonhowerVerasViei\Downloads\1859\2. RKO Alimentos\Anexo 32 - Extrato Fundo Titânia.pdf"),
    Path(r"C:\Users\AlysonhowerVerasViei\Downloads\1859\2. RKO Alimentos\Anexo 36 - despacho-dicol-assinado.pdf")
]

async def load_cookies(tab):
    """Load saved cookies to restore login session."""
    if not COOKIE_FILE.exists():
        return False
    
    try:
        saved_cookies = json.loads(COOKIE_FILE.read_text(encoding='utf-8'))
        
        # Convert to simplified format (only settable fields)
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


def parse_copilot_response(text: str, filename: str = '') -> dict:
    """Extract contents from <data>, <resumo>, and <objeto> tags."""
    tags = ['data', 'resumo', 'objeto']
    result = {'Correspondência': filename}
    
    for tag in tags:
        pattern = rf'<{tag}>(.*?)</{tag}>'
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        result[tag.capitalize()] = match.group(1).strip() if match else ''
    
    return result


def append_to_excel(data: dict, filepath: Path):
    """Append data to Excel file, creating it with headers if it doesn't exist."""
    headers = list(data.keys())
    
    if filepath.exists():
        wb = load_workbook(filepath)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
    
    ws.append([data[h] for h in headers])
    wb.save(filepath)



async def interact_and_send(tab, file_path: Path):
    """
    Performs the interaction up to clicking 'Send'.
    Returns True if successful, False otherwise.
    """
    filename = file_path.name
    print(f"Starting interaction for: {filename}")

    try:
        # Navigate if needed (on new tabs or first run)
        # Check if we are already on the correct page to avoid unnecessary reloads if possible,
        # but for safety/consistency, ensuring the URL is correct is good.
        current_url = await tab.current_url
        if "m365.cloud.microsoft/chat" not in current_url:
            await tab.go_to('https://m365.cloud.microsoft/chat')
            await asyncio.sleep(random.uniform(2.0, 4.0))

        # === FILE ATTACHMENT ===
        # Click the plus menu button
        plus_menu_btn = await tab.find(data_testid="PlusMenuButton", timeout=10)
        await plus_menu_btn.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-5, 5),
            hold_time=random.uniform(0.08, 0.15)
        )

        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Attach file
        async with tab.expect_file_chooser(files=[file_path]):
            upload_menu_item = await tab.find(text="Carregar imagens e arquivos", timeout=5)
            await upload_menu_item.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-3, 3),
                hold_time=random.uniform(0.08, 0.15)
            )

        # Wait for file processing
        await asyncio.sleep(random.uniform(1.0, 2.0))

        # Focus chat
        chat_input = await tab.find(aria_label="Copilot de Mensagens", timeout=10)
        await chat_input.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15)
        )
        
        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Type message
        message = r"Analise o documento em anexo e execute as três tarefas a seguir: 1. Identifique a data do documento (se existir), no formato DD/MM/AAAA (dia/mês/ano). Se não existir deixe em branco; 2. Produza um resumo em um único parágrafo, claro e objetivo; 3. Em seguida, identifique e destaque o objeto central do documento em uma única sentença curta, no estilo punchline (poucas palavras, direto ao ponto, refletindo o cerne do conteúdo). Retorne estritamente neste formato: `<data>[Data do documento ou vazio se não existir]</data><resumo>[Resumo em um único parágrafo]</resumo><objeto>[Objeto central do documento]</objeto>"
        await chat_input.type_text(message, humanize=True)

        await asyncio.sleep(random.uniform(0.5, 1.5))

        # Click Send
        send_button = await tab.find(aria_label="Enviar", timeout=5)
        await send_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.09, 0.18)
        )
        
        print(f"Sent request for {filename}.")
        return True

    except Exception as e:
        print(f"Error during interaction for {filename}: {e}")
        return False


async def wait_and_save(tab, file_path: Path):
    """
    Waits for the response in an already active tab and saves it.
    """
    filename = file_path.name
    print(f"Waiting for response for: {filename} in background...")

    try:
        # Wait for valid response (Copy Button)
        copy_button = await tab.find(
            data_testid="CopyButtonTestId",
            timeout=180 
        )

        # Human-like delay before copying
        await asyncio.sleep(random.uniform(0.5, 1.0))

        await copy_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15)
        )

        await asyncio.sleep(random.uniform(0.3, 0.5))

        # Retrieve text
        clipboard_result = await tab.execute_script(
            "return navigator.clipboard.readText()", 
            await_promise=True
        )
        
        try:
            response_text = clipboard_result['result']['result']['value']
        except (KeyError, TypeError) as e:
            print(f"Warning: Unexpected clipboard structure for {filename}: {e}")
            response_text = ""

        if response_text:
            attachment_name = file_path.stem
            parsed_data = parse_copilot_response(response_text, filename=attachment_name)
            
            output_file = Path(__file__).parent / 'outputs' / 'copilot_responses.xlsx'
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            async with EXCEL_LOCK:
                append_to_excel(parsed_data, output_file)
                print(f"✅ Results for {filename} saved to {output_file}")
        else:
            print(f"❌ No response content obtained for {filename}")

    except Exception as e:
        print(f"Error waiting for response for {filename}: {e}")


async def copilot_chat_automation():
    """Main automation flow with sequential initiation and parallel waiting."""
    options = ChromiumOptions()
    options.block_notifications = True
    options.block_popups = True

    async with Chrome(options=options) as browser:
        # Initialize browser and cookies
        first_tab = await browser.start()
        
        try:
            await browser.grant_permissions(
                permissions=['clipboardReadWrite', 'clipboardSanitizedWrite'],
                origin='https://m365.cloud.microsoft'
            )
        except Exception as e:
            print(f"Warning: Could not grant permissions: {e}")

        await load_cookies(first_tab)
        
        waiting_tasks = []
        
        for i, file_path in enumerate(FILES_TO_PROCESS):
            if not file_path.exists():
                print(f"Error: File not found: {file_path}")
                continue

            # Determine which tab to use
            if i == 0:
                tab = first_tab
            else:
                # Open new tab for subsequent files
                print(f"Opening new tab for {file_path.name}...")
                tab = await browser.new_tab()
            
            # 1. SEQUENTIAL PART: Interact and Send
            # We await this, so the next file only starts processing after this one is sent.
            success = await interact_and_send(tab, file_path)
            
            if success:
                # 2. PARALLEL PART: Wait for response
                # We schedule this task but do not await it immediately.
                # It runs in the background while the main loop initiates the next file.
                task = asyncio.create_task(wait_and_save(tab, file_path))
                waiting_tasks.append(task)
            else:
                print(f"Skipping wait for {file_path.name} due to send failure.")
                if tab != first_tab:
                    await tab.close()

        # After all files have been initiated, wait for all responses to complete
        if waiting_tasks:
            print("All prompts sent. Waiting for all responses to complete...")
            await asyncio.gather(*waiting_tasks)

        await save_cookies(browser)
        input("Processamento concluído. Pressione Enter para fechar o navegador...")

if __name__ == "__main__":
    asyncio.run(copilot_chat_automation())