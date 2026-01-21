"""
Unit tests for the scheduler module.
Tests time parsing, schedule window checking, and edge cases.
"""

import pytest
from datetime import time

from scheduler import parse_time, is_within_schedule


class TestParseTime:
    """Tests for parse_time function."""
    
    def test_parse_time_valid_morning(self):
        """Parses morning time correctly."""
        result = parse_time("07:00")
        assert result == time(7, 0)
    
    def test_parse_time_valid_evening(self):
        """Parses evening time correctly."""
        result = parse_time("20:30")
        assert result == time(20, 30)
    
    def test_parse_time_midnight(self):
        """Parses midnight correctly."""
        result = parse_time("00:00")
        assert result == time(0, 0)
    
    def test_parse_time_with_spaces(self):
        """Handles whitespace."""
        result = parse_time("  09:15  ")
        assert result == time(9, 15)
    
    def test_parse_time_invalid_format(self):
        """Raises ValueError for bad format."""
        with pytest.raises(ValueError):
            parse_time("7:00:00")
    
    def test_parse_time_invalid_separator(self):
        """Raises ValueError for wrong separator."""
        with pytest.raises(ValueError):
            parse_time("07-00")
    
    def test_parse_time_invalid_hour(self):
        """Raises ValueError for invalid hour."""
        with pytest.raises(ValueError):
            parse_time("25:00")
    
    def test_parse_time_invalid_minute(self):
        """Raises ValueError for invalid minute."""
        with pytest.raises(ValueError):
            parse_time("12:60")


class TestIsWithinSchedule:
    """Tests for is_within_schedule function."""
    
    # Normal schedule tests (start < stop)
    def test_within_schedule_normal(self):
        """10:00 is within 09:00-17:00."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(10, 0)
        assert is_within_schedule(start, stop, now) is True
    
    def test_outside_schedule_before(self):
        """06:00 is outside 09:00-17:00."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(6, 0)
        assert is_within_schedule(start, stop, now) is False
    
    def test_outside_schedule_after(self):
        """22:00 is outside 09:00-17:00."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(22, 0)
        assert is_within_schedule(start, stop, now) is False
    
    # Overnight schedule tests (start > stop)
    def test_overnight_schedule_within_night(self):
        """23:00 is within 22:00-06:00."""
        start = time(22, 0)
        stop = time(6, 0)
        now = time(23, 0)
        assert is_within_schedule(start, stop, now) is True
    
    def test_overnight_schedule_within_morning(self):
        """03:00 is within 22:00-06:00."""
        start = time(22, 0)
        stop = time(6, 0)
        now = time(3, 0)
        assert is_within_schedule(start, stop, now) is True
    
    def test_overnight_schedule_outside(self):
        """12:00 is outside 22:00-06:00."""
        start = time(22, 0)
        stop = time(6, 0)
        now = time(12, 0)
        assert is_within_schedule(start, stop, now) is False
    
    # Boundary condition tests
    def test_boundary_exactly_at_start(self):
        """Exactly at start time = within schedule."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(9, 0)
        assert is_within_schedule(start, stop, now) is True
    
    def test_boundary_exactly_at_stop(self):
        """Exactly at stop time = outside schedule."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(17, 0)
        assert is_within_schedule(start, stop, now) is False
    
    def test_boundary_one_minute_before_stop(self):
        """One minute before stop = within schedule."""
        start = time(9, 0)
        stop = time(17, 0)
        now = time(16, 59)
        assert is_within_schedule(start, stop, now) is True


class TestCalculateBackoffDelay:
    """Tests for calculate_backoff_delay function."""
    
    def test_first_attempt(self):
        """First attempt (0) returns initial delay."""
        from scheduler import calculate_backoff_delay
        assert calculate_backoff_delay(0, initial_delay=60) == 60
    
    def test_second_attempt(self):
        """Second attempt doubles the delay."""
        from scheduler import calculate_backoff_delay
        assert calculate_backoff_delay(1, initial_delay=60) == 120
    
    def test_third_attempt(self):
        """Third attempt quadruples the initial delay."""
        from scheduler import calculate_backoff_delay
        assert calculate_backoff_delay(2, initial_delay=60) == 240
    
    def test_max_delay_cap(self):
        """Delay is capped at max_delay (24 hours by default)."""
        from scheduler import calculate_backoff_delay
        # After many attempts, should not exceed max
        result = calculate_backoff_delay(20, initial_delay=60, max_delay=86400)
        assert result == 86400
    
    def test_custom_backoff_factor(self):
        """Custom backoff factor works."""
        from scheduler import calculate_backoff_delay
        # Factor of 3: 60 -> 180 -> 540
        assert calculate_backoff_delay(1, initial_delay=60, backoff_factor=3) == 180


class TestFormatDuration:
    """Tests for format_duration function."""
    
    def test_seconds(self):
        """Formats seconds correctly."""
        from scheduler import format_duration
        assert format_duration(30) == "30s"
    
    def test_minutes(self):
        """Formats minutes correctly."""
        from scheduler import format_duration
        assert format_duration(300) == "5 min"
    
    def test_hours(self):
        """Formats hours correctly."""
        from scheduler import format_duration
        assert format_duration(7200) == "2h"
    
    def test_hours_and_minutes(self):
        """Formats hours and minutes correctly."""
        from scheduler import format_duration
        assert format_duration(5400) == "1h 30min"
