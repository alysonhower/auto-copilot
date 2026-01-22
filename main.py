import asyncio
import json
import random
import re
from datetime import time
from pathlib import Path
from typing import List, Optional, Tuple

import click
import win32com.client
import pythoncom

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions

from scheduler import (
    parse_time,
    is_within_schedule,
    wait_until_schedule_starts,
    retry_with_backoff,
)

# Arquivo persistente de cookies
COOKIE_FILE = Path(__file__).parent / ".cookies.json"

# Lock removed - COM handles concurrency


# Extensões suportadas
SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".md",
    ".txt",
    ".pdf",
    ".docx",
    ".jpg",
    ".jpeg",
    ".png",
}


async def load_cookies(tab):
    """Carrega cookies salvos para restaurar a sessão de login."""
    if not COOKIE_FILE.exists():
        return False

    try:
        saved_cookies = json.loads(COOKIE_FILE.read_text(encoding="utf-8"))

        simplified_cookies = []
        for cookie in saved_cookies:
            simplified = {
                "name": cookie["name"],
                "value": cookie["value"],
                "domain": cookie.get("domain"),
                "path": cookie.get("path", "/"),
                "secure": cookie.get("secure", False),
                "httpOnly": cookie.get("httpOnly", False),
            }
            if "expires" in cookie and cookie["expires"] > 0:
                simplified["expires"] = cookie["expires"]
            simplified_cookies.append(simplified)

        await tab.set_cookies(simplified_cookies)
        click.echo(f"Carregados {len(simplified_cookies)} cookies de {COOKIE_FILE}")
        return True
    except Exception as e:
        click.echo(f"Erro ao carregar cookies: {e}", err=True)
        return False


async def save_cookies(browser):
    """Salva cookies após login bem-sucedido."""
    try:
        cookies = await browser.get_cookies()
        COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding="utf-8")
        click.echo(f"Salvos {len(cookies)} cookies em {COOKIE_FILE}")
    except Exception as e:
        click.echo(f"Erro ao salvar cookies: {e}", err=True)


def load_message_content(message_input: str) -> str:
    """
    Carrega o conteúdo da mensagem a partir de um arquivo .md ou retorna a string diretamente.

    Args:
        message_input: Caminho para arquivo .md ou string de mensagem direta

    Returns:
        String com o conteúdo da mensagem
    """
    # Verifica se é um caminho de arquivo
    message_path = Path(message_input)

    if message_path.exists() and message_path.is_file():
        if message_path.suffix.lower() == ".md":
            try:
                content = message_path.read_text(encoding="utf-8")
                click.echo(f"📄 Mensagem carregada de: {message_path}")
                return content.strip()
            except Exception as e:
                click.echo(
                    f"Erro ao ler arquivo de mensagem {message_path}: {e}", err=True
                )
                raise
        else:
            click.echo(
                f"Aviso: Arquivo {message_path} não é .md, usando caminho como mensagem literal",
                err=True,
            )

    # Se não for arquivo válido, retorna como string literal
    return message_input


def parse_copilot_response(text: str, filename: str = "") -> dict:
    """Extrai conteúdo das tags <data>, <resumo>, e <objeto>."""
    tags = ["data", "resumo", "objeto"]
    result = {"Correspondência": filename}

    for tag in tags:
        pattern = rf"<{tag}>(.*?)</{tag}>"
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        result[tag.capitalize()] = match.group(1).strip() if match else ""

    return result


def get_excel_app():
    """Obtém ou cria uma instância do Excel."""
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        try:
            return win32com.client.Dispatch("Excel.Application")
        except Exception as e:
            click.echo(f"Erro ao inicializar Excel: {e}", err=True)
            return None


def get_processed_filenames(filepath: Path) -> set:
    """Retorna o conjunto de nomes de arquivos já processados no Excel usando COM."""
    if not filepath.exists():
        return set()

    pythoncom.CoInitialize()
    app = get_excel_app()
    if not app:
        return set()

    wb = None
    opened_by_us = False
    abs_path = str(filepath.resolve())
    processed = set()

    try:
        # Tenta encontrar workbook já aberto
        try:
            # Workbooks collection é 1-indexed? Em python itera objects.
            # Acesso direto pode falhar se busy.
            for w in app.Workbooks:
                if w.FullName.lower() == abs_path.lower():
                    wb = w
                    break
        except Exception:
            pass  # Pode falhar se Excel estiver ocupado

        if not wb:
            try:
                wb = app.Workbooks.Open(abs_path)
                opened_by_us = True
            except Exception as e:
                click.echo(f"Erro ao abrir arquivo {filepath}: {e}", err=True)
                return set()

        ws = wb.Worksheets(1)

        # Ler headers
        headers = []
        col = 1
        while True:
            val = ws.Cells(1, col).Value
            if not val:
                break
            headers.append(val)
            col += 1

        if "Correspondência" in headers:
            col_idx = headers.index("Correspondência") + 1
            # Ler coluna (assume dados contíguos ou varre used range)
            # UsedRange é mais seguro
            used_range = ws.UsedRange
            # Convert values to list of lists
            data = used_range.Value
            if data and isinstance(data, tuple):
                # data é tuple de tuples
                # Row 1 é header.
                for i, row in enumerate(data):
                    if i == 0:
                        continue  # Skip header
                    if len(row) >= col_idx:
                        val = row[col_idx - 1]
                        if val:
                            processed.add(str(val))
    except Exception as e:
        click.echo(f"Erro ao ler Excel via COM: {e}", err=True)
    finally:
        if opened_by_us and wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    return processed


def append_to_excel(data: dict, filepath: Path):
    """Adiciona dados ao Excel usando COM com repetição em caso de bloqueio."""
    pythoncom.CoInitialize()
    abs_path = str(filepath.resolve())

    max_retries = 20
    for attempt in range(max_retries):
        try:
            app = get_excel_app()
            if not app:
                return

            wb = None
            opened_by_us = False

            # Check open workbooks
            try:
                for w in app.Workbooks:
                    if w.FullName.lower() == abs_path.lower():
                        wb = w
                        break
            except Exception:
                pass

            if not wb:
                if filepath.exists():
                    wb = app.Workbooks.Open(abs_path)
                    opened_by_us = True
                else:
                    wb = app.Workbooks.Add()
                    wb.SaveAs(abs_path)
                    opened_by_us = True

            ws = wb.Worksheets(1)

            # Find next row
            # xlUp = -4162
            last_row_cell = ws.Cells(ws.Rows.Count, 1).End(-4162)
            last_row = last_row_cell.Row

            # Check empty sheet
            if last_row == 1 and not ws.Cells(1, 1).Value:
                next_row = 1
            else:
                next_row = last_row + 1

            # Map headers
            sheet_headers = []
            col = 1
            while True:
                val = ws.Cells(1, col).Value
                if not val:
                    break
                sheet_headers.append(str(val))
                col += 1

            if not sheet_headers:
                # Initialize new sheet headers
                headers = list(data.keys())
                for i, h in enumerate(headers, 1):
                    ws.Cells(1, i).Value = h
                sheet_headers = headers
                next_row = 2

            # Write data matching headers
            for key, value in data.items():
                if key in sheet_headers:
                    col_idx = sheet_headers.index(key) + 1
                    ws.Cells(next_row, col_idx).Value = value
                else:
                    # New column? Append logic could go here but skipping for simplicity
                    pass

            # Save logic: Only save/close if we opened it.
            # If user has it open, changes appear live.
            if opened_by_us:
                wb.Save()
                wb.Close()

            return  # Success

        except Exception as e:
            # Handle CallRejected (user typing)
            err_str = str(e)
            if "Call was rejected by callee" in err_str or "-2147418111" in err_str:
                click.echo(
                    f"⚠️ Excel ocupado (usuário editando?), tentativa {attempt + 1}/{max_retries}...",
                    err=True,
                )
                time.sleep(2)
                continue
            else:
                click.echo(f"Erro fatal ao escrever no Excel: {e}", err=True)
                if opened_by_us and wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                raise e


async def safe_close_tab(tab, timeout: float = 5.0):
    """
    Safely attempts to close a tab with a short timeout.
    Does not raise exceptions - used for cleanup scenarios.

    Args:
        tab: The tab to close
        timeout: Maximum time to wait for close operation (default: 5s)
    """
    try:
        await asyncio.wait_for(tab.close(), timeout=timeout)
    except asyncio.TimeoutError:
        click.echo("Aviso: Timeout ao fechar aba (ignorando)", err=True)
    except Exception as e:
        click.echo(f"Aviso: Não foi possível fechar aba: {e}", err=True)


async def interact_and_send(tab, file_path: Path, message: str):
    """
    Realiza a interação até clicar em 'Enviar'.
    Retorna True se bem-sucedido, False caso contrário.
    """
    filename = file_path.name
    click.echo(f"Iniciando interação para: {filename}")

    try:
        # Navega se necessário
        current_url = await tab.current_url
        if "m365.cloud.microsoft/chat" not in current_url:
            await tab.go_to("https://m365.cloud.microsoft/chat")
            await asyncio.sleep(random.uniform(2.0, 4.0))

        # Simulate user reading/exploring the page
        for _ in range(random.randint(1, 2)):
            await tab.scroll.by("down", random.randint(100, 300), humanize=True)
            await asyncio.sleep(random.uniform(0.5, 1.0))
        await tab.scroll.to_top()
        await asyncio.sleep(random.uniform(0.3, 0.6))

        # === SELECIONAR CHAT TEMPORÁRIO ===
        try:
            # 1. Clicar no accordion "Chat temporário"
            chat_temp_accordion = await tab.find(
                data_automation_id="newPrivateChatMenuButton", timeout=30
            )
            await chat_temp_accordion.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.08, 0.15),
            )
            await asyncio.sleep(random.uniform(0.5, 1.0))

            # 2. Clicar no botão "Chat temporário"
            chat_temp_button = await tab.find(
                data_automation_id="newPrivateChatButton", timeout=30
            )
            await chat_temp_button.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.08, 0.15),
            )
            await asyncio.sleep(random.uniform(1.0, 2.0))  # Wait for new chat context

        except Exception as e:
            click.echo(f"Erro ao selecionar Chat temporário: {e}", err=True)
            # Opsional: Raise se for crítico ou continuar tentando no chat padrão?
            # Se a UI não abrir, provavelmente falhará adiante, então raise.
            raise

        # === ANEXAR ARQUIVO ===
        try:
            plus_menu_btn = await tab.find(data_testid="PlusMenuButton", timeout=60)
            await plus_menu_btn.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.08, 0.15),
            )
        except Exception:
            # Tenta recuperar se o menu não abrir ou já estiver aberto?
            # Por simplicidade, assume erro se não achar.
            # Se falhar aqui, pode ser que a página não carregou direito.
            raise

        await asyncio.sleep(random.uniform(0.3, 0.8))

        async with tab.expect_file_chooser(files=[file_path]):
            upload_menu_item = await tab.find(
                text="Carregar imagens e arquivos", timeout=5
            )
            await upload_menu_item.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-3, 3),
                hold_time=random.uniform(0.08, 0.15),
            )

        await asyncio.sleep(random.uniform(1.0, 2.0))

        # Focar chat
        chat_input = await tab.find(aria_label="Copilot de Mensagens", timeout=60)
        await chat_input.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15),
        )

        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Digitar mensagem (recebida como parâmetro)
        await chat_input.type_text(message, humanize=True)

        # User reviewing message before sending (longer pause)
        await asyncio.sleep(random.uniform(1.5, 3.0))

        # Clicar Enviar
        send_button = await tab.find(aria_label="Enviar", timeout=60)
        await send_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.09, 0.18),
        )

        # User watching AI confirmation that generation started
        await asyncio.sleep(random.uniform(1.5, 3.0))

        click.echo(f"Solicitação enviada para {filename}.")
        return True

    except Exception as e:
        click.echo(f"Erro durante interação para {filename}: {e}", err=True)
        # Re-raise to allow retry_with_backoff to handle it
        raise


async def wait_and_save(tab, file_path: Path, output_file: Path):
    """
    Aguarda a resposta em uma aba já ativa e salva no Excel.
    """
    filename = file_path.name
    click.echo(f"Aguardando resposta para: {filename} em segundo plano...")

    try:
        # Aguarda botão de copiar (indica fim da geração)
        copy_button = await tab.find(data_testid="CopyButtonTestId", timeout=180)

        await asyncio.sleep(random.uniform(0.5, 1.0))

        await copy_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15),
        )

        await asyncio.sleep(random.uniform(0.3, 0.5))

        # Ler clipboard
        clipboard_result = await tab.execute_script(
            "return navigator.clipboard.readText()", await_promise=True
        )

        try:
            response_text = clipboard_result["result"]["result"]["value"]
        except (KeyError, TypeError) as e:
            click.echo(
                f"Aviso: Estrutura inesperada do clipboard para {filename}: {e}",
                err=True,
            )
            response_text = ""

        if response_text:
            attachment_name = file_path.stem
            parsed_data = parse_copilot_response(
                response_text, filename=attachment_name
            )

            output_file.parent.mkdir(parents=True, exist_ok=True)

            # async with EXCEL_LOCK:
            # Lock removed for COM
            append_to_excel(parsed_data, output_file)
            click.echo(f"✅ Resultados para {filename} salvos em {output_file}")

        else:
            click.echo(
                f"❌ Nenhum conteúdo obtido na resposta para {filename}", err=True
            )

    except Exception as e:
        click.echo(f"Erro aguardando resposta para {filename}: {e}", err=True)


def resolve_paths(paths: Tuple[str]) -> List[Path]:
    """
    Resolve caminhos fornecidos (arquivos e diretórios) para uma lista de arquivos.
    Busca recursivamente em diretórios por extensões suportadas.
    """
    files_to_process = []

    for path_str in paths:
        path = Path(path_str)
        if not path.exists():
            click.echo(f"Aviso: Caminho não encontrado ignorado: {path}", err=True)
            continue

        if path.is_file():
            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                files_to_process.append(path)
            else:
                click.echo(
                    f"Aviso: Extensão não suportada ignorada: {path.name}", err=True
                )

        elif path.is_dir():
            click.echo(f"Escaneando diretório: {path}")
            for item in path.rglob("*"):
                if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                    files_to_process.append(item)

    return files_to_process


async def process_files_logic(
    files: List[Path],
    output_file: Path,
    message: str,
    start_time: Optional[time] = None,
    stop_time: Optional[time] = None,
    disable_headless: bool = False,
):
    """Lógica principal de orquestração do navegador."""
    if not files:
        click.echo("Nenhum arquivo válido encontrado para processar.")
        return

    # Verifica arquivos já processados para retomada
    processed_filenames = get_processed_filenames(output_file)

    if processed_filenames:
        click.echo(
            f"📋 Encontrados {len(processed_filenames)} arquivos já processados no Excel."
        )

    # Filtra arquivos que ainda precisam ser processados
    pending_files = [f for f in files if f.stem not in processed_filenames]
    skipped_count = len(files) - len(pending_files)

    if skipped_count > 0:
        click.echo(f"⏭️  Pulando {skipped_count} arquivos já processados.")

    if not pending_files:
        click.echo("✅ Todos os arquivos já foram processados. Nada a fazer.")
        return

    click.echo(f"📁 {len(pending_files)} arquivos pendentes para processamento.")

    options = ChromiumOptions()
    options.block_notifications = True
    options.block_popups = True

    # WebGL (software renderer to avoid unique GPU signatures)
    options.add_argument("--use-gl=swiftshader")
    options.add_argument("--disable-features=WebGLDraftExtensions")

    # WebRTC IP leak prevention
    options.add_argument("--force-webrtc-ip-handling-policy=disable_non_proxied_udp")

    # Headless mode (enabled by default, use --disable-headless to show browser)
    if not disable_headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    async with Chrome(options=options) as browser:
        first_tab = await browser.start()

        try:
            await browser.grant_permissions(
                permissions=["clipboardReadWrite", "clipboardSanitizedWrite"],
                origin="https://m365.cloud.microsoft",
            )
        except Exception as e:
            click.echo(f"Aviso: Não foi possível conceder permissões: {e}", err=True)

        await load_cookies(first_tab)

        waiting_tasks = []
        first_success = False  # Track if we've had at least one success

        click.echo(f"Iniciando processamento de {len(pending_files)} arquivos...")

        for i, file_path in enumerate(pending_files):
            # Check schedule before each file
            if start_time and stop_time:
                if not is_within_schedule(start_time, stop_time):
                    await wait_until_schedule_starts(start_time, stop_time)

            # Determina qual aba usar
            if i == 0:
                tab = first_tab
            else:
                click.echo(f"Abrindo nova aba para {file_path.name}...")
                tab = await browser.new_tab()

            # 1. PARTE SEQUENCIAL: Interagir e Enviar
            try:
                if not first_success:
                    # Before first success: unlimited retry with backoff
                    # (system might not be ready, e.g. Copilot not released yet)
                    await retry_with_backoff(
                        interact_and_send,
                        tab,
                        file_path,
                        message,
                        max_retries=None,  # Unlimited until success
                        start_time=start_time,
                        stop_time=stop_time,
                    )
                else:
                    # After first success: no retry, skip on error
                    # (likely file-specific issue)
                    await interact_and_send(tab, file_path, message)

                success = True
                first_success = True
            except Exception as e:
                if not first_success:
                    # Should only get here if schedule ended
                    click.echo(f"Falha definitiva para {file_path.name}: {e}", err=True)
                else:
                    click.echo(f"Erro com {file_path.name}, pulando: {e}", err=True)
                success = False

            if success:
                # 2. PARTE PARALELA: Aguardar resposta
                task = asyncio.create_task(wait_and_save(tab, file_path, output_file))
                waiting_tasks.append((task, tab))
            else:
                if tab != first_tab:
                    await safe_close_tab(tab)

            # === GERENCIAMENTO DE MEMÓRIA (FECHAR ABAS ANTIGAS) ===
            # Se atingir 10 abas abertas aguardando, fecha as 5 mais antigas.
            if len(waiting_tasks) >= 10:
                click.echo(
                    "⚠️ Limite de 10 abas atingido. Fechando as 5 mais antigas para liberar memória..."
                )

                # Seleciona as 5 tarefas mais antigas (primeiros 5 da lista)
                tasks_to_close = waiting_tasks[:5]
                waiting_tasks = waiting_tasks[5:]

                # Separa tasks e tabs
                tasks_only = [t for t, _ in tasks_to_close]
                tabs_to_close = [tab_ for _, tab_ in tasks_to_close]

                # Aguarda tasks terminarem
                await asyncio.gather(*tasks_only, return_exceptions=True)

                # Fecha as abas (simulando humano fechando uma por uma)
                for old_tab in tabs_to_close:
                    if old_tab != first_tab:
                        await safe_close_tab(old_tab)
                        # Pequena pausa entre fechamentos para parecer natural
                        await asyncio.sleep(random.uniform(0.3, 0.7))

                click.echo("✅ Limpeza de memória concluída (5 abas fechadas).")

        if waiting_tasks:
            click.echo("Todas as solicitações enviadas. Aguardando respostas...")
            tasks_only = [t for t, _ in waiting_tasks]
            await asyncio.gather(*tasks_only, return_exceptions=True)

        # Graceful cleanup: close all tabs before browser exit
        click.echo("Limpando abas...")
        all_tabs = [tab for _, tab in waiting_tasks]
        if first_tab not in all_tabs:
            all_tabs.append(first_tab)

        for tab in all_tabs:
            await safe_close_tab(tab)

        # Small delay to let browser process pending operations
        await asyncio.sleep(0.5)

        await save_cookies(browser)


@click.command()
@click.option(
    "--prompt",
    "-p",
    required=True,
    type=str,
    help="Mensagem a ser enviada ao Copilot ou caminho para arquivo .md contendo a mensagem",
)
@click.argument("paths", nargs=-1, type=click.Path(exists=True))
@click.option(
    "--output",
    "-o",
    type=click.Path(dir_okay=False),
    default="outputs/copilot_responses.xlsx",
    help="Caminho do arquivo Excel de saída (padrão: outputs/copilot_responses.xlsx)",
)
@click.option(
    "--start",
    "-st",
    type=str,
    default=None,
    help="Hora de início do processamento (HH:MM, ex: 07:00)",
)
@click.option(
    "--stop",
    "-stp",
    type=str,
    default=None,
    help="Hora de término do processamento (HH:MM, ex: 20:20)",
)
@click.option(
    "--disable-headless",
    "-dh",
    is_flag=True,
    default=False,
    help="Desabilitar modo headless (mostrar o navegador)",
)
def main(prompt, paths, output, start, stop, disable_headless):
    """
    Auto-Copilot CLI.

    Processa arquivos usando o Microsoft 365 Copilot e salva os resultados em Excel.

    Exemplos de uso:
        auto-copilot -p path/to/prompt.md path/to/dir path/to/file1 path/to/file2

    PATHS: Caminhos para arquivos ou diretórios a serem processados.
           Diretórios são escaneados recursivamente por arquivos suportados
           (.xlsx, .xls, .md, .txt, .pdf, .docx, .jpg, .jpeg, .png).
    """
    # Load message content (from file or direct string)
    try:
        message_content = load_message_content(prompt)
    except Exception as e:
        click.echo(f"Erro ao carregar mensagem: {e}", err=True)
        return

    # Validate schedule options
    parsed_start = None
    parsed_stop = None

    if start and stop:
        try:
            parsed_start = parse_time(start)
            parsed_stop = parse_time(stop)
            click.echo(
                f"⏰ Agendamento ativo: {parsed_start.strftime('%H:%M')} - {parsed_stop.strftime('%H:%M')}"
            )
        except ValueError as e:
            click.echo(f"Erro: {e}", err=True)
            return
    elif start or stop:
        click.echo("Erro: --start e --stop devem ser usados juntos", err=True)
        return

    files = resolve_paths(paths)

    if not files:
        click.echo("Por favor, forneça pelo menos um arquivo ou diretório válido.")
        return

    # Resolve o caminho do arquivo de saída
    output_file = Path(output)
    if not output_file.is_absolute():
        output_file = Path(__file__).parent / output_file

    click.echo(f"📊 Arquivo de saída: {output_file}")

    # Executa o loop assíncrono
    asyncio.run(
        process_files_logic(
            files,
            output_file,
            message_content,
            parsed_start,
            parsed_stop,
            disable_headless,
        )
    )


if __name__ == "__main__":
    main()
