import asyncio
import json
import random
import re
import time
from datetime import time as dt_time
from pathlib import Path
from typing import List, Optional, Tuple

import click
import win32com.client
from selectolax.parser import HTMLParser

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions
from pydoll.constants import Key
from pydoll.protocol.browser.types import PermissionType

from scheduler import (
    parse_time,
    is_within_schedule,
    wait_until_schedule_starts,
    retry_with_backoff,
)
from clipboard_utils import safe_set_clipboard, safe_get_clipboard

COOKIE_FILE = Path(__file__).parent / ".cookies.json"

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


class UploadFailedError(Exception):
    """Raised when file upload fails (Try again button detected)."""

    pass


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
        click.echo(
            click.style(
                f"Carregados {len(simplified_cookies)} cookies de {COOKIE_FILE}",
                fg="green",
            )
        )
        return True
    except Exception as e:
        click.echo(click.style(f"Erro ao carregar cookies: {e}", fg="red"), err=True)
        return False


async def save_cookies(browser):
    """Salva cookies após login bem-sucedido."""
    try:
        cookies = await browser.get_cookies()
        COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding="utf-8")
        click.echo(
            click.style(f"Salvos {len(cookies)} cookies em {COOKIE_FILE}", fg="green")
        )
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
                click.echo(
                    click.style(f"Mensagem carregada de: {message_path}", fg="cyan")
                )
                return content.strip()
            except Exception as e:
                click.echo(
                    click.style(
                        f"Erro ao ler arquivo de mensagem {message_path}: {e}", fg="red"
                    ),
                    err=True,
                )
                raise
        else:
            click.echo(
                click.style(
                    f"Aviso: Arquivo {message_path} não é .md, usando caminho como mensagem literal",
                    fg="yellow",
                ),
                err=True,
            )

    # Se não for arquivo válido, retorna como string literal
    return message_input


def normalize_column_name(tag: str) -> str:
    """
    Converte uma tag (ex: 'my-tag', 'myTag') em um nome de coluna legível (ex: 'My Tag').
    """
    # Substitui hífens e underscores por espaços
    s = tag.replace("-", " ").replace("_", " ")

    # Insere espaço antes de letras maiúsculas (camelCase)
    s = re.sub(r"(?<!^)(?=[A-Z])", " ", s)

    # Capitaliza cada palavra
    return s.title()


def extract_tags_from_prompt(prompt_text: str) -> List[str]:
    """
    Analisa o prompt para descobrir quais tags o usuário espera na resposta.
    Usa Selectolax para percorrer a estrutura.
    """
    tags = []
    seen = set()

    # Define tags padrão HTML para ignorar se aparecerem sem intencionalidade clara
    # (Embora num prompt markdown, qualquer tag <foo> seja relevante)
    ignored_tags = {"html", "head", "body", "-text", "br", "p", "div", "span"}

    try:
        tree = HTMLParser(prompt_text)

        # Percorre todos os nós
        if tree.root:
            for node in tree.root.traverse():
                tag_name = node.tag

                if not tag_name or not isinstance(tag_name, str):
                    continue

                # Filtra tags irrelevantes e nomes estranhos
                if tag_name in ignored_tags:
                    continue

                if tag_name not in seen:
                    tags.append(tag_name)
                    seen.add(tag_name)

    except Exception as e:
        click.echo(f"Erro ao extrair tags do prompt: {e}", err=True)

    return tags


def parse_copilot_response(
    text: str, tags: List[str], filename: str = "", name_column: str = "Arquivo"
) -> dict:
    """
    Extrai conteúdo das tags especificadas usando Selectolax.

    Args:
        text: O texto HTML/Markdown da resposta.
        tags: Lista de tags para buscar (em ordem).
        filename: Nome do arquivo processado (para metadados).
        name_column: Nome da coluna onde o nome do arquivo será salvo.
    """
    result = {name_column: filename}

    try:
        tree = HTMLParser(text)

        for tag in tags:
            column_name = normalize_column_name(tag)

            # Busca o primeiro nó com a tag
            node = tree.css_first(tag)
            if node:
                # Extrai texto limpo
                content = node.text(strip=True)
                result[column_name] = content
            else:
                result[column_name] = ""

    except Exception as e:
        click.echo(f"Erro ao parsear resposta com Selectolax: {e}", err=True)
        # Pode retornar parcial ou vazio

    return result


def get_excel_app():
    """Obtém ou cria uma instância do Excel."""
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        try:
            return win32com.client.Dispatch("Excel.Application")
        except Exception as e:
            click.echo(
                click.style(f"Erro ao inicializar Excel: {e}", fg="red"), err=True
            )
            return None


def get_processed_filenames(filepath: Path, name_column: str = "Arquivo") -> set:
    """Retorna o conjunto de nomes de arquivos já processados no Excel usando COM."""
    if not filepath.exists():
        return set()

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
                click.echo(
                    click.style(f"Erro ao abrir arquivo {filepath}: {e}", fg="red"),
                    err=True,
                )
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

        if name_column in headers:
            col_idx = headers.index(name_column) + 1
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
        click.echo(click.style(f"Erro ao ler Excel via COM: {e}", fg="red"), err=True)
    finally:
        if opened_by_us and wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    return processed


def format_excel_table(ws):
    """
    Formata a planilha como uma Tabela Oficial do Excel e aplica estilos.
    """
    try:
        # Constantes Excel
        xlSrcRange = 1
        xlYes = 1
        xlCenter = -4108
        xlVAlignCenter = -4108
        xlContinuous = 1
        xlThin = 2

        used_range = ws.UsedRange

        # Ignora se planilha vazia
        if used_range.Count == 1 and not used_range.Value:
            try:
                # Tenta verificar se a unica celula tem valor (Count 1 as vezes engana)
                if not ws.Cells(1, 1).Value:
                    return
            except Exception:
                return

        # 1. Cria ou Atualiza Tabela (ListObject)
        tbl = None
        if ws.ListObjects.Count == 0:
            try:
                tbl = ws.ListObjects.Add(xlSrcRange, used_range, None, xlYes)
                tbl.Name = "CopilotData"
                tbl.ShowAutoFilter = True
            except Exception as e:
                click.echo(f"Aviso ao criar tabela: {e}", err=True)
        else:
            try:
                tbl = ws.ListObjects(1)
                tbl.Resize(used_range)
            except Exception:
                pass

        # Aplica estilo neutro
        if tbl:
            try:
                tbl.TableStyle = "TableStyleLight9"
            except Exception:
                pass

        # 2. Formatação de Fonte e Alinhamento
        # Aplica em TODO o UsedRange (Headers + Data)
        used_range.Font.Name = "Calibri"
        used_range.Font.Size = 11
        used_range.HorizontalAlignment = xlCenter
        used_range.VerticalAlignment = xlVAlignCenter
        used_range.WrapText = True

        # 3. Bordas (All Borders)
        try:
            used_range.Borders.LineStyle = xlContinuous
            used_range.Borders.Weight = xlThin
        except Exception:
            pass

        # 4. AutoFit Colunas
        # AutoFit as vezes deixa colunas muito largas para textos longos (ex: corpo do email)
        # Mas é o pedido: "proper headers... formatting etc"
        used_range.Columns.AutoFit()

    except Exception as e:
        click.echo(f"Erro não crítico ao formatar tabela: {e}", err=True)


def append_to_excel(data: dict, filepath: Path):
    """Adiciona dados ao Excel usando COM com repetição em caso de bloqueio."""
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

            # Aplica formatação de tabela
            format_excel_table(ws)

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
                    f"Excel ocupado (usuário editando?), tentativa {attempt + 1}/{max_retries}...",
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
        click.echo(
            click.style("Timeout ao fechar aba (ignorando)", fg="yellow"), err=True
        )
    except Exception as e:
        click.echo(click.style(f"Não foi possível fechar aba: {e}", fg="red"), err=True)


async def interact_and_send(
    browser, tab, file_path: Path, message: str, risky_mode: bool = False
):
    """
    Realiza a interação até clicar em 'Enviar'.
    Retorna a aba ativa (pode mudar em caso de retry).
    """
    filename = file_path.name

    # Retry loop for upload failures
    for attempt in range(3):
        try:
            click.echo(
                click.style(
                    f"Iniciando interação ({filename}) - Tentativa {attempt + 1}/3...",
                    fg="cyan",
                )
            )

            # Navega se necessário
            current_url = await tab.current_url

            if "m365.cloud.microsoft/chat" not in current_url:
                await tab.go_to("https://m365.cloud.microsoft/chat")
                await asyncio.sleep(random.uniform(0.5, 3.0))

            # === SELECIONAR CHAT TEMPORÁRIO ===
            try:
                # 1. Clicar no accordion "Chat temporário"
                chat_temp_accordion = await tab.find(
                    data_automation_id="newPrivateChatMenuButton", timeout=60
                )

                click.echo(
                    click.style(f"Abrindo Chat temporário ({filename})...", fg="cyan")
                )

                await chat_temp_accordion.click(
                    x_offset=random.randint(-5, 5),
                    y_offset=random.randint(-5, 5),
                    hold_time=random.uniform(0.02, 0.15),
                )
                await asyncio.sleep(random.uniform(0.1, 1.0))

                # 2. Clicar no botão "Chat temporário"
                chat_temp_button = await tab.find(
                    data_automation_id="newPrivateChatButton", timeout=60
                )

                await chat_temp_button.click(
                    x_offset=random.randint(-5, 5),
                    y_offset=random.randint(-5, 5),
                    hold_time=random.uniform(0.02, 0.15),
                )

                await asyncio.sleep(
                    random.uniform(0.1, 1.0)
                )  # Wait for new chat context

            except Exception as e:
                click.echo(
                    click.style(
                        f"Erro ao selecionar Chat temporário ({filename}): {e}",
                        fg="red",
                    ),
                    err=True,
                )
                raise

            # === ANEXAR ARQUIVO ===
            click.echo(click.style(f"Abrindo menu de anexos ({filename})", fg="cyan"))

            try:
                plus_menu_btn = await tab.find(data_testid="PlusMenuButton", timeout=60)
                await plus_menu_btn.click(
                    x_offset=random.randint(-5, 5),
                    y_offset=random.randint(-5, 5),
                    hold_time=random.uniform(0.02, 0.15),
                )
            except Exception:
                raise

            await asyncio.sleep(random.uniform(0.1, 1.0))

            async with tab.expect_file_chooser(files=[file_path]):
                upload_menu_item = await tab.find(
                    text="Carregar imagens e arquivos", timeout=60
                )
                await upload_menu_item.click(
                    x_offset=random.randint(-5, 5),
                    y_offset=random.randint(-5, 5),
                    hold_time=random.uniform(0.02, 0.15),
                )

            await asyncio.sleep(random.uniform(1.0, 2.0))

            # === CHECK FOR UPLOAD FAILURE ===
            # Check for "Tentar novamente" button (quietly, without logging timeout error)
            try_again_btn = await tab.find(
                text="Tentar novamente", timeout=5.0, raise_exc=False
            )
            if try_again_btn:
                raise UploadFailedError("Botão 'Tentar novamente' detectado.")

            # Focar chat
            chat_input = await tab.find(aria_label="Copilot de Mensagens", timeout=60)

            await chat_input.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.02, 0.15),
            )

            await asyncio.sleep(random.uniform(0.1, 1.0))

            click.echo(click.style(f"Digitando prompt ({filename})...", fg="cyan"))

            # Digitar mensagem
            if risky_mode:
                safe_set_clipboard(message)
                await tab.keyboard.hotkey(Key.CONTROL, Key.V)
            else:
                await chat_input.type_text(message, humanize=True)

            await asyncio.sleep(random.uniform(1.5, 3.0))

            send_button = await tab.find(aria_label="Enviar", timeout=60)

            await send_button.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.02, 0.15),
            )

            click.echo(
                click.style(f"Aguardando início da geração ({filename})...", fg="cyan")
            )
            try:
                await tab.find(aria_label="Interromper geração", timeout=60)
                click.echo(click.style(f"Geração iniciada ({filename})...", fg="green"))
            except Exception:
                click.echo(
                    click.style(
                        f"Botão de parar não encontrado, assumindo geração iniciada ({filename})...",
                        fg="yellow",
                    )
                )

            await asyncio.sleep(random.uniform(0.5, 3.0))

            return tab

        except UploadFailedError as e:
            click.echo(
                click.style(
                    f"Falha de upload ({e}). Tentando novamente imediatamente...",
                    fg="magenta",
                ),
                err=True,
            )
            await safe_close_tab(tab)
            await asyncio.sleep(random.uniform(1.0, 2.0))
            tab = await browser.new_tab()
            continue

        except Exception as e:
            click.echo(click.style(f"{e}", fg="red"), err=True)
            raise

    raise Exception(f"Falha no upload após 3 tentativas para {filename}")


async def wait_and_save(
    tab,
    file_path: Path,
    output_file: Path,
    tags: List[str],
    name_column: str = "Arquivo",
):
    """
    Aguarda a resposta em uma aba já ativa e salva no Excel.
    """
    filename = file_path.name
    click.echo(
        click.style(f"Aguardando resposta ({filename}) em segundo plano...", fg="cyan")
    )

    try:
        # Aguarda botão de copiar (indica fim da geração)
        copy_button = await tab.find(aria_label="Copiar Resposta", timeout=180)

        await copy_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15),
        )

        response_text = safe_get_clipboard()

        if response_text:
            attachment_name = file_path.stem
            parsed_data = parse_copilot_response(
                response_text,
                tags=tags,
                filename=attachment_name,
                name_column=name_column,
            )

            output_file.parent.mkdir(parents=True, exist_ok=True)

            append_to_excel(parsed_data, output_file)
            click.echo(
                click.style(
                    f"Resultados para {filename} salvos em {output_file}", fg="green"
                )
            )

        else:
            click.echo(
                click.style(
                    f"Nenhum conteúdo obtido na resposta ({filename})", fg="red"
                ),
                err=True,
            )

    except Exception as e:
        click.echo(f"{e}", err=True)


def resolve_paths(paths: Tuple[str]) -> List[Path]:
    """
    Resolve caminhos fornecidos (arquivos e diretórios) para uma lista de arquivos.
    Busca recursivamente em diretórios por extensões suportadas.
    """
    files_to_process = []

    for path_str in paths:
        path = Path(path_str)
        if not path.exists():
            click.echo(
                click.style(f"Caminho não encontrado ignorado ({path})", fg="yellow"),
                err=True,
            )
            continue

        if path.is_file():
            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                files_to_process.append(path)
            else:
                click.echo(
                    f"Extensão de arquivo não suportada ignorada ({path.name})",
                    err=True,
                )

        elif path.is_dir():
            click.echo(click.style(f"Escaneando diretório ({path})...", fg="cyan"))
            for item in path.rglob("*"):
                if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                    files_to_process.append(item)

    return files_to_process


async def process_files_logic(
    files: List[Path],
    output_file: Path,
    message: str,
    tags: List[str],
    start_time: Optional[dt_time] = None,
    stop_time: Optional[dt_time] = None,
    disable_headless: bool = False,
    risky_mode: bool = False,
    name_column: str = "Arquivo",
):
    """Lógica principal de orquestração do navegador."""
    if not files:
        click.echo(
            click.style("Nenhum arquivo válido encontrado para processar.", fg="red")
        )
        return

    # Verifica arquivos já processados para retomada
    processed_filenames = get_processed_filenames(output_file, name_column=name_column)

    if processed_filenames:
        click.echo(
            click.style(
                f"Encontrados {len(processed_filenames)} arquivos já processados no Excel.",
                fg="yellow",
            )
        )

    # Filtra arquivos que ainda precisam ser processados
    pending_files = [f for f in files if f.stem not in processed_filenames]
    skipped_count = len(files) - len(pending_files)

    if skipped_count > 0:
        click.echo(
            click.style(
                f"Pulando {skipped_count} arquivos já processados.", fg="yellow"
            )
        )

    if not pending_files:
        click.echo(
            click.style(
                "Todos os arquivos já foram processados. Nada a fazer.", fg="green"
            )
        )
        return

    click.echo(
        click.style(
            f"{len(pending_files)} arquivos pendentes para processamento.", fg="cyan"
        )
    )

    options = ChromiumOptions()
    options.block_notifications = True
    options.block_popups = True

    # Core stealth
    # options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")

    # WebGL (software renderer to avoid unique GPU signatures)
    options.add_argument("--use-gl=swiftshader")
    options.add_argument("--disable-features=WebGLDraftExtensions")

    # WebRTC IP leak prevention
    options.add_argument("--force-webrtc-ip-handling-policy=disable_non_proxied_udp")

    # Permissions and first-run
    # options.add_argument("--no-first-run") # Already added by Pydoll by default
    # options.add_argument("--no-default-browser-check") # Already added by Pydoll by default

    # Disable unnecessary features
    options.add_argument("--disable-translate")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")

    # Network optimizations
    options.add_argument("--disable-features=NetworkPrediction")
    options.add_argument("--dns-prefetch-disable")

    # Headless mode (enabled by default, use --disable-headless to show browser)
    if not disable_headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    async with Chrome(options=options) as browser:
        first_tab = await browser.start()

        try:
            await browser.grant_permissions(
                permissions=[
                    PermissionType.CLIPBOARD_READ_WRITE,
                    PermissionType.CLIPBOARD_SANITIZED_WRITE,
                ],
                origin="https://m365.cloud.microsoft",
            )
        except Exception as e:
            click.echo(f"Não foi possível conceder permissões: {e}", err=True)

        await load_cookies(first_tab)

        waiting_tasks = []
        first_success = False  # Track if we've had at least one success

        click.echo(
            click.style(
                f"Iniciando processamento de {len(pending_files)} arquivos...",
                fg="green",
            )
        )

        for i, file_path in enumerate(pending_files):
            # Check schedule before each file
            if start_time and stop_time:
                if not is_within_schedule(start_time, stop_time):
                    await wait_until_schedule_starts(start_time, stop_time)

            # Determina qual aba usar
            if i == 0:
                tab = first_tab
            else:
                click.echo(
                    click.style(f"Abrindo nova aba para {file_path.name}...", fg="cyan")
                )
                tab = await browser.new_tab()

            # 1. PARTE SEQUENCIAL: Interagir e Enviar
            try:
                if not first_success:
                    # Before first success: unlimited retry with backoff
                    # (system might not be ready, e.g. Copilot not released yet)
                    tab = await retry_with_backoff(
                        interact_and_send,
                        browser,  # Pass browser
                        tab,
                        file_path,
                        message,
                        risky_mode,  # Pass risky_mode
                        max_retries=None,  # Unlimited until success
                        start_time=start_time,
                        stop_time=stop_time,
                    )
                else:
                    # After first success: no retry, skip on error
                    # (likely file-specific issue)
                    tab = await interact_and_send(
                        browser, tab, file_path, message, risky_mode=risky_mode
                    )

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
                task = asyncio.create_task(
                    wait_and_save(tab, file_path, output_file, tags, name_column)
                )
                waiting_tasks.append((task, tab))
            else:
                if tab != first_tab:
                    await safe_close_tab(tab)

            # === GERENCIAMENTO DE MEMÓRIA (FECHAR ABAS ANTIGAS) ===
            # Se atingir 10 abas abertas aguardando, fecha as 5 mais antigas.
            if len(waiting_tasks) >= 10:
                click.echo(
                    "Limite de 10 abas atingido. Fechando as 5 mais antigas para liberar memória..."
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
                        await asyncio.sleep(random.uniform(1.0, 2.0))

                click.echo("Limpeza de memória concluída (5 abas fechadas).")

        if waiting_tasks:
            click.echo("Todas as solicitações enviadas. Aguardando respostas...")
            tasks_only = [t for t, _ in waiting_tasks]
            await asyncio.gather(*tasks_only, return_exceptions=True)

        await save_cookies(browser)

        click.echo("Limpando abas...")

        all_tabs = [tab for _, tab in waiting_tasks]
        if first_tab not in all_tabs:
            all_tabs.append(first_tab)

        for tab in all_tabs:
            await safe_close_tab(tab)

        await asyncio.sleep(0.5)


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
@click.option(
    "--name-column",
    "-nc",
    type=str,
    default="Arquivo",
    help="Nome da coluna para os nomes dos arquivos (padrão: Arquivo)",
)
@click.option(
    "--risky",
    is_flag=True,
    default=False,
    help="Habilita modo arriscado (copiar/colar prompt) para maior velocidade",
)
def main(prompt, paths, output, start, stop, disable_headless, name_column, risky):
    """
    Auto-Copilot CLI.

    Processa arquivos usando o Microsoft 365 Copilot e salva os resultados em Excel.

    Exemplos de uso:
        auto-copilot -p path/to/prompt.md path/to/dir path/to/file1 path/to/file2

    PATHS: Caminhos para arquivos ou diretórios a serem processados.
           Diretórios são escaneados recursivamente por arquivos suportados
           (.xlsx, .xls, .md, .txt, .pdf, .docx, .jpg, .jpeg, .png).
    """
    if risky:
        click.echo(
            click.style(
                "⚠️  MODO ARRISCADO HABILITADO  ⚠️", fg="yellow", bold=True, blink=True
            )
        )
        click.echo(
            click.style(
                "Isso fará com que o prompt seja colado (Ctrl+V) em vez de digitado.",
                fg="yellow",
            )
        )
        click.echo(
            click.style(
                "Isso torna o processo mais rápido, mas o comportamento é mais parecido com o de um robô.",
                fg="yellow",
            )
        )
        click.echo("")

        confirm = click.prompt(
            "Deseja continuar? [Sim/sim/s]", default="Não", show_default=True
        )
        if confirm.lower() not in ["sim", "s"]:
            click.echo(click.style("Operação cancelada pelo usuário.", fg="red"))
            return
        click.echo(
            click.style("Modo arriscado confirmado. Prosseguindo...", fg="green")
        )

    # Capitalize the column name
    name_column = name_column.title()

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
                f"Agendamento ativo: {parsed_start.strftime('%H:%M')} - {parsed_stop.strftime('%H:%M')}"
            )
        except ValueError as e:
            click.echo(f"{e}", err=True)
            return
    elif start or stop:
        click.echo("--start e --stop devem ser usados juntos", err=True)
        return

    files = resolve_paths(paths)

    if not files:
        click.echo("Por favor, forneça pelo menos um arquivo ou diretório válido.")
        return

    # Resolve o caminho do arquivo de saída
    output_file = Path(output)
    if not output_file.is_absolute():
        output_file = Path(__file__).parent / output_file

    click.echo(f"Arquivo de saída: {output_file}")

    # Extract tags from prompt content to determine dynamic columns
    click.echo("Analisando prompt para identificar tags dinâmicas...")
    tags = extract_tags_from_prompt(message_content)

    if tags:
        click.echo(f"Tags identificadas: {', '.join(tags)}")
    else:
        click.echo("Nenhuma tag específica identificada no prompt.")

    # Executa o loop assíncrono
    asyncio.run(
        process_files_logic(
            files,
            output_file,
            message_content,
            tags,
            parsed_start,
            parsed_stop,
            disable_headless,
            risky,
            name_column=name_column,
        )
    )


if __name__ == "__main__":
    main()
