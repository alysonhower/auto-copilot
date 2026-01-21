"""
Scheduling module for programmable time-window processing.

Allows the app to only process files within user-specified hours,
pausing automatically outside that window.
"""

import asyncio
from datetime import time, datetime
from typing import Optional

import click


def parse_time(time_str: str) -> time:
    """
    Parse time string in HH:MM format to time object.

    Args:
        time_str: Time string in HH:MM format (e.g., "07:00", "20:30")

    Returns:
        time object

    Raises:
        ValueError: If format is invalid
    """
    try:
        parts = time_str.strip().split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid time format: {time_str}")

        hour = int(parts[0])
        minute = int(parts[1])

        if not (0 <= hour <= 23) or not (0 <= minute <= 59):
            raise ValueError(f"Invalid time values: {time_str}")

        return time(hour, minute)
    except (ValueError, AttributeError) as e:
        raise ValueError(
            f"Invalid time format '{time_str}'. Use HH:MM (e.g., 07:00)"
        ) from e


def is_within_schedule(
    start_time: time, stop_time: time, now: Optional[time] = None
) -> bool:
    """
    Check if current time is within the active schedule window.

    Handles both normal schedules (e.g., 09:00-17:00) and overnight schedules
    (e.g., 22:00-06:00 where stop < start).

    Args:
        start_time: When processing should start
        stop_time: When processing should stop
        now: Current time (defaults to datetime.now().time())

    Returns:
        True if within schedule, False otherwise

    Note:
        - Start time is inclusive (exactly at start = within schedule)
        - Stop time is exclusive (exactly at stop = outside schedule)
    """
    if now is None:
        now = datetime.now().time()

    # Normal schedule: start < stop (e.g., 09:00 to 17:00)
    if start_time <= stop_time:
        return start_time <= now < stop_time

    # Overnight schedule: start > stop (e.g., 22:00 to 06:00)
    # Within schedule if: now >= start OR now < stop
    return now >= start_time or now < stop_time


async def wait_until_schedule_starts(start_time: time, stop_time: time) -> None:
    """
    If currently outside the schedule, wait until the start time.

    Sleeps in short intervals to remain responsive to signals (e.g., Ctrl+C).
    Prints status messages to keep user informed.

    Args:
        start_time: When processing should start
        stop_time: When processing should stop
    """
    CHECK_INTERVAL = 30  # seconds between checks

    if is_within_schedule(start_time, stop_time):
        return

    click.echo(
        f"⏸️  Outside schedule ({start_time.strftime('%H:%M')} - {stop_time.strftime('%H:%M')})"
    )
    click.echo(f"⏳ Waiting until {start_time.strftime('%H:%M')} to resume...")

    while not is_within_schedule(start_time, stop_time):
        await asyncio.sleep(CHECK_INTERVAL)

    click.echo("▶️  Schedule window opened. Resuming processing...")


# Retry constants
DEFAULT_INITIAL_DELAY = 60  # 1 minute
DEFAULT_MAX_DELAY = 86400  # 24 hours (can skip a day if needed)
DEFAULT_BACKOFF_FACTOR = 2  # Double each time


def calculate_backoff_delay(
    attempt: int,
    initial_delay: int = DEFAULT_INITIAL_DELAY,
    max_delay: int = DEFAULT_MAX_DELAY,
    backoff_factor: float = DEFAULT_BACKOFF_FACTOR,
) -> int:
    """
    Calculate delay for exponential backoff.

    Args:
        attempt: Current attempt number (0-indexed)
        initial_delay: Initial delay in seconds (default: 60s = 1 min)
        max_delay: Maximum delay in seconds (default: 7200s = 2 hours)
        backoff_factor: Multiplier for each attempt (default: 2)

    Returns:
        Delay in seconds, capped at max_delay
    """
    delay = initial_delay * (backoff_factor**attempt)
    return min(int(delay), max_delay)


def format_duration(seconds: int) -> str:
    """Format seconds into human-readable duration (e.g., '5 min', '1h 30min')."""
    if seconds < 60:
        return f"{seconds}s"
    elif seconds < 3600:
        minutes = seconds // 60
        return f"{minutes} min"
    else:
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        if minutes:
            return f"{hours}h {minutes}min"
        return f"{hours}h"


async def retry_with_backoff(
    coro_func,
    *args,
    max_retries: int = None,
    initial_delay: int = DEFAULT_INITIAL_DELAY,
    max_delay: int = DEFAULT_MAX_DELAY,
    start_time: Optional[time] = None,
    stop_time: Optional[time] = None,
    **kwargs,
):
    """
    Execute an async function with exponential backoff retry on failure.

    Args:
        coro_func: Async function to execute
        *args: Arguments to pass to the function
        max_retries: Maximum retry attempts (None = unlimited while in schedule)
        initial_delay: Initial delay in seconds (default: 60s)
        max_delay: Maximum delay in seconds (default: 7200s = 2 hours)
        start_time: Schedule start time (for schedule-aware retries)
        stop_time: Schedule stop time (for schedule-aware retries)
        **kwargs: Keyword arguments to pass to the function

    Returns:
        Result from the async function on success

    Raises:
        Exception: Re-raises the last exception if all retries exhausted
    """
    attempt = 0

    while True:
        try:
            result = await coro_func(*args, **kwargs)
            if attempt > 0:
                click.echo(f"✅ Succeeded after {attempt} retries")
            return result
        except Exception as e:
            # Check if max retries reached
            if max_retries is not None and attempt >= max_retries:
                click.echo(
                    f"❌ Max retries ({max_retries}) reached. Giving up.", err=True
                )
                raise

            # Check if outside schedule (should wait for schedule instead of retry)
            if (
                start_time
                and stop_time
                and not is_within_schedule(start_time, stop_time)
            ):
                click.echo(
                    "⏸️  Outside schedule. Will retry when schedule resumes.", err=True
                )
                await wait_until_schedule_starts(start_time, stop_time)
                attempt = 0  # Reset attempts after schedule wait
                continue

            delay = calculate_backoff_delay(attempt, initial_delay, max_delay)
            click.echo(f"⚠️  Attempt {attempt + 1} failed: {e}", err=True)
            click.echo(f"🔄 Retrying in {format_duration(delay)}...")

            await asyncio.sleep(delay)
            attempt += 1
