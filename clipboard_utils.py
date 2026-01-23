import time
import win32clipboard
import win32con
import pywintypes
import click


def safe_set_clipboard(text: str, retries: int = 20, delay: float = 0.5):
    """
    Sets text to the clipboard safely, handling "Access Denied" errors with retries.
    Uses win32clipboard (pywin32) for robust Windows clipboard handling.

    Args:
        text (str): The text to copy to the clipboard.
        retries (int): Number of retries if the clipboard is locked.
        delay (float): Seconds to wait between retries.
    """
    last_error = None

    for attempt in range(retries):
        try:
            # Open the clipboard
            win32clipboard.OpenClipboard()
            try:
                # Empty it (required before setting)
                win32clipboard.EmptyClipboard()
                # Set text (using CF_UNICODETEXT for modern string support)
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            finally:
                # Always close specifically in this block to ensure we release lock
                win32clipboard.CloseClipboard()

            # If we get here, success
            return

        except pywintypes.error as e:
            # Error 5 is Access Denied (clipboard locked by another app)
            # Error 1418 is "Thread does not have a clipboard open" (shouldn't happen inside correct logic but valid to catch)
            last_error = e
            if e.winerror == 5:  # ACCESS_DENIED
                time.sleep(delay)
                continue
            else:
                # Other errors, maybe re-raise or log and retry?
                # Let's retry for robustness but log it
                time.sleep(delay)
                continue
        except Exception as e:
            last_error = e
            time.sleep(delay)
            continue

    # If we exhausted retries
    click.echo(
        click.style(
            f"Falha ao definir clipboard após {retries} tentativas: {last_error}",
            fg="red",
        ),
        err=True,
    )
    raise last_error


def safe_get_clipboard(retries: int = 20, delay: float = 0.5) -> str:
    """
    Gets text from the clipboard safely, handling concurrency.

    Args:
        retries (int): Number of retries if the clipboard is locked.
        delay (float): Seconds to wait between retries.

    Returns:
        str: The text content of the clipboard.
    """
    last_error = None

    for attempt in range(retries):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    return data
                elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                    data = win32clipboard.GetClipboardData(win32con.CF_TEXT)
                    return data.decode("utf-8", errors="ignore")
                else:
                    return ""  # No text data available
            finally:
                win32clipboard.CloseClipboard()

        except pywintypes.error as e:
            last_error = e
            if e.winerror == 5:  # ACCESS_DENIED
                time.sleep(delay)
                continue
            else:
                time.sleep(delay)
                continue
        except Exception as e:
            last_error = e
            time.sleep(delay)
            continue

    click.echo(
        click.style(
            f"Falha ao ler clipboard após {retries} tentativas: {last_error}", fg="red"
        ),
        err=True,
    )
    return ""
