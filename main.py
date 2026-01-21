import asyncio
import json
import random
import re
from datetime import time
from pathlib import Path
from typing import List, Optional, Tuple

import click
from openpyxl import Workbook, load_workbook

from pydoll.browser.chromium import Chrome
from pydoll.browser.options import ChromiumOptions

from scheduler import (
    parse_time,
    is_within_schedule,
    wait_until_schedule_starts,
    retry_with_backoff
)

# Arquivo persistente de cookies
COOKIE_FILE = Path(__file__).parent / '.cookies.json'

# Lock global para escrita no Excel
EXCEL_LOCK = asyncio.Lock()

# Extensões suportadas
SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.md', '.txt', '.pdf', '.docx', '.jpg', '.jpeg', '.png'}

async def load_cookies(tab):
    """Carrega cookies salvos para restaurar a sessão de login."""
    if not COOKIE_FILE.exists():
        return False
    
    try:
        saved_cookies = json.loads(COOKIE_FILE.read_text(encoding='utf-8'))
        
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
        click.echo(f"Carregados {len(simplified_cookies)} cookies de {COOKIE_FILE}")
        return True
    except Exception as e:
        click.echo(f"Erro ao carregar cookies: {e}", err=True)
        return False


async def save_cookies(browser):
    """Salva cookies após login bem-sucedido."""
    try:
        cookies = await browser.get_cookies()
        COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding='utf-8')
        click.echo(f"Salvos {len(cookies)} cookies em {COOKIE_FILE}")
    except Exception as e:
        click.echo(f"Erro ao salvar cookies: {e}", err=True)


def parse_copilot_response(text: str, filename: str = '') -> dict:
    """Extrai conteúdo das tags <data>, <resumo>, e <objeto>."""
    tags = ['data', 'resumo', 'objeto']
    result = {'Correspondência': filename}
    
    for tag in tags:
        pattern = rf'<{tag}>(.*?)</{tag}>'
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        result[tag.capitalize()] = match.group(1).strip() if match else ''
    
    return result


def get_processed_filenames(filepath: Path) -> set:
    """Retorna o conjunto de nomes de arquivos já processados no Excel."""
    if not filepath.exists():
        return set()
    
    try:
        wb = load_workbook(filepath, read_only=True)
        ws = wb.active
        
        # Encontra o índice da coluna 'Correspondência'
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        if 'Correspondência' not in header_row:
            return set()
        
        col_idx = header_row.index('Correspondência')
        
        # Coleta todos os nomes de arquivos (pula o header)
        processed = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) > col_idx and row[col_idx]:
                processed.add(row[col_idx])
        
        wb.close()
        return processed
    except Exception as e:
        click.echo(f"Aviso: Erro ao ler arquivos já processados: {e}", err=True)
        return set()


def append_to_excel(data: dict, filepath: Path):
    """Adiciona dados ao arquivo Excel, criando-o se não existir."""
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
    Realiza a interação até clicar em 'Enviar'.
    Retorna True se bem-sucedido, False caso contrário.
    """
    filename = file_path.name
    click.echo(f"Iniciando interação para: {filename}")

    try:
        # Navega se necessário
        current_url = await tab.current_url
        if "m365.cloud.microsoft/chat" not in current_url:
            await tab.go_to('https://m365.cloud.microsoft/chat')
            await asyncio.sleep(random.uniform(2.0, 4.0))

        # Simulate user reading/exploring the page
        for _ in range(random.randint(1, 2)):
            await tab.scroll.by('down', random.randint(100, 300), humanize=True)
            await asyncio.sleep(random.uniform(0.5, 1.0))
        await tab.scroll.to_top()
        await asyncio.sleep(random.uniform(0.3, 0.6))

        # === ANEXAR ARQUIVO ===
        try:
            plus_menu_btn = await tab.find(data_testid="PlusMenuButton", timeout=10)
            await plus_menu_btn.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-5, 5),
                hold_time=random.uniform(0.08, 0.15)
            )
        except Exception:
             # Tenta recuperar se o menu não abrir ou já estiver aberto?
             # Por simplicidade, assume erro se não achar.
             # Se falhar aqui, pode ser que a página não carregou direito.
             raise

        await asyncio.sleep(random.uniform(0.3, 0.8))

        async with tab.expect_file_chooser(files=[file_path]):
            upload_menu_item = await tab.find(text="Carregar imagens e arquivos", timeout=5)
            await upload_menu_item.click(
                x_offset=random.randint(-5, 5),
                y_offset=random.randint(-3, 3),
                hold_time=random.uniform(0.08, 0.15)
            )

        await asyncio.sleep(random.uniform(1.0, 2.0))

        # Focar chat
        chat_input = await tab.find(aria_label="Copilot de Mensagens", timeout=10)
        await chat_input.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15)
        )
        
        await asyncio.sleep(random.uniform(0.3, 0.8))

        # Digitar mensagem
        message = r"Analise o documento em anexo e execute as três tarefas a seguir: 1. Identifique a data do documento (se existir), no formato DD/MM/AAAA (dia/mês/ano). Se não existir deixe em branco; 2. Produza um resumo em um único parágrafo, claro e objetivo; 3. Em seguida, identifique e destaque o objeto central do documento em uma única sentença curta, no estilo punchline (poucas palavras, direto ao ponto, refletindo o cerne do conteúdo). IMPORTANTE: Escreva de forma direta, sem usar frases de meta-referência como 'Este documento', 'O documento apresenta', 'O arquivo trata de', etc. Comunique as informações diretamente, como se estivesse relatando os fatos sem mencionar que é um documento. Retorne estritamente neste formato: `<data>[Data do documento ou vazio se não existir]</data><resumo>[Resumo direto, sem meta-referências]</resumo><objeto>[Objeto central, de forma direta]</objeto>"
        await chat_input.type_text(message, humanize=True)

        # User reviewing message before sending (longer pause)
        await asyncio.sleep(random.uniform(1.5, 3.0))

        # Clicar Enviar
        send_button = await tab.find(aria_label="Enviar", timeout=5)
        await send_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.09, 0.18)
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
        copy_button = await tab.find(
            data_testid="CopyButtonTestId",
            timeout=180 
        )

        await asyncio.sleep(random.uniform(0.5, 1.0))

        await copy_button.click(
            x_offset=random.randint(-5, 5),
            y_offset=random.randint(-3, 3),
            hold_time=random.uniform(0.08, 0.15)
        )

        await asyncio.sleep(random.uniform(0.3, 0.5))

        # Ler clipboard
        clipboard_result = await tab.execute_script(
            "return navigator.clipboard.readText()", 
            await_promise=True
        )
        
        try:
            response_text = clipboard_result['result']['result']['value']
        except (KeyError, TypeError) as e:
            click.echo(f"Aviso: Estrutura inesperada do clipboard para {filename}: {e}", err=True)
            response_text = ""

        if response_text:
            attachment_name = file_path.stem
            parsed_data = parse_copilot_response(response_text, filename=attachment_name)
            
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            async with EXCEL_LOCK:
                append_to_excel(parsed_data, output_file)
                click.echo(f"✅ Resultados para {filename} salvos em {output_file}")
        else:
            click.echo(f"❌ Nenhum conteúdo obtido na resposta para {filename}", err=True)

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
                click.echo(f"Aviso: Extensão não suportada ignorada: {path.name}", err=True)
        
        elif path.is_dir():
            click.echo(f"Escaneando diretório: {path}")
            for item in path.rglob('*'):
                if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                    files_to_process.append(item)
    
    return files_to_process


async def process_files_logic(
    files: List[Path],
    output_file: Path,
    start_time: Optional[time] = None,
    stop_time: Optional[time] = None
):
    """Lógica principal de orquestração do navegador."""
    if not files:
        click.echo("Nenhum arquivo válido encontrado para processar.")
        return

    # Verifica arquivos já processados para retomada
    processed_filenames = get_processed_filenames(output_file)
    
    if processed_filenames:
        click.echo(f"📋 Encontrados {len(processed_filenames)} arquivos já processados no Excel.")
    
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
    options.add_argument('--use-gl=swiftshader')
    options.add_argument('--disable-features=WebGLDraftExtensions')

    # WebRTC IP leak prevention
    options.add_argument('--force-webrtc-ip-handling-policy=disable_non_proxied_udp')

    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")

    async with Chrome(options=options) as browser:
        first_tab = await browser.start()
        
        try:
            await browser.grant_permissions(
                permissions=['clipboardReadWrite', 'clipboardSanitizedWrite'],
                origin='https://m365.cloud.microsoft'
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
                        max_retries=None,  # Unlimited until success
                        start_time=start_time,
                        stop_time=stop_time
                    )
                else:
                    # After first success: no retry, skip on error
                    # (likely file-specific issue)
                    await interact_and_send(tab, file_path)
                
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
                waiting_tasks.append(task)
            else:
                if tab != first_tab:
                    await tab.close()

        if waiting_tasks:
            click.echo("Todas as solicitações enviadas. Aguardando respostas...")
            await asyncio.gather(*waiting_tasks)

        await save_cookies(browser)
        click.echo("Pressione Enter para fechar o navegador...")
        input()


@click.command()
@click.argument('paths', nargs=-1, type=click.Path(exists=True))
@click.option(
    '--output', '-o',
    type=click.Path(dir_okay=False),
    default='outputs/copilot_responses.xlsx',
    help='Caminho do arquivo Excel de saída (padrão: outputs/copilot_responses.xlsx)'
)
@click.option(
    '--start',
    type=str,
    default=None,
    help='Hora de início do processamento (HH:MM, ex: 07:00)'
)
@click.option(
    '--stop',
    type=str,
    default=None,
    help='Hora de término do processamento (HH:MM, ex: 20:20)'
)
def main(paths, output, start, stop):
    """
    Auto-Copilot CLI.
    
    Processa arquivos usando o Microsoft 365 Copilot e salva os resultados em Excel.
    
    PATHS: Caminhos para arquivos ou diretórios a serem processados.
           Diretórios são escaneados recursivamente por arquivos suportados
           (.xlsx, .xls, .md, .txt, .pdf, .docx, .jpg, .jpeg, .png).
    """
    # Validate schedule options
    parsed_start = None
    parsed_stop = None
    
    if start and stop:
        try:
            parsed_start = parse_time(start)
            parsed_stop = parse_time(stop)
            click.echo(f"⏰ Agendamento ativo: {parsed_start.strftime('%H:%M')} - {parsed_stop.strftime('%H:%M')}")
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
    asyncio.run(process_files_logic(files, output_file, parsed_start, parsed_stop))


if __name__ == "__main__":
    main()