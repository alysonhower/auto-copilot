"""
Tests for clipboard lock functionality to prevent race conditions
when multiple tabs operate on the shared clipboard concurrently.
"""

import asyncio
import pytest
from unittest.mock import AsyncMock

# Import the module under test
from main import (
    get_clipboard_lock,
    clipboard_write_and_paste,
    clipboard_click_copy_and_read,
)


class TestClipboardLock:
    """Tests for the clipboard lock mechanism."""

    def setup_method(self):
        """Reset the global lock before each test."""
        import main
        main.CLIPBOARD_LOCK = None

    def test_get_clipboard_lock_creates_lock(self):
        """Test that get_clipboard_lock creates a new lock if none exists."""
        import main
        main.CLIPBOARD_LOCK = None
        
        lock = get_clipboard_lock()
        
        assert lock is not None
        assert isinstance(lock, asyncio.Lock)

    def test_get_clipboard_lock_returns_same_lock(self):
        """Test that get_clipboard_lock returns the same lock on subsequent calls."""
        import main
        main.CLIPBOARD_LOCK = None
        
        lock1 = get_clipboard_lock()
        lock2 = get_clipboard_lock()
        
        assert lock1 is lock2

    async def test_clipboard_write_and_paste_acquires_lock(self):
        """Test that clipboard_write_and_paste acquires the lock during operation."""
        import main
        main.CLIPBOARD_LOCK = None
        
        # Create mock tab
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock()
        mock_tab.keyboard = AsyncMock()
        mock_tab.keyboard.hotkey = AsyncMock()
        
        lock = get_clipboard_lock()
        lock_acquired_during_operation = False
        
        original_execute_script = mock_tab.execute_script
        
        async def check_lock_during_execute(*args, **kwargs):
            nonlocal lock_acquired_during_operation
            lock_acquired_during_operation = lock.locked()
            return await original_execute_script(*args, **kwargs)
        
        mock_tab.execute_script = check_lock_during_execute
        
        await clipboard_write_and_paste(mock_tab, "test text", "test.txt")
        
        assert lock_acquired_during_operation, "Lock should be held during clipboard operation"

    async def test_clipboard_click_copy_and_read_acquires_lock(self):
        """Test that clipboard_click_copy_and_read acquires the lock during operation."""
        import main
        main.CLIPBOARD_LOCK = None
        
        # Create mock tab and copy button
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock(return_value={
            "result": {"result": {"value": "copied text"}}
        })
        
        mock_copy_button = AsyncMock()
        mock_copy_button.click = AsyncMock()
        
        lock = get_clipboard_lock()
        lock_acquired_during_operation = False
        
        original_click = mock_copy_button.click
        
        async def check_lock_during_click(*args, **kwargs):
            nonlocal lock_acquired_during_operation
            lock_acquired_during_operation = lock.locked()
            return await original_click(*args, **kwargs)
        
        mock_copy_button.click = check_lock_during_click
        
        result = await clipboard_click_copy_and_read(mock_tab, mock_copy_button, "test.txt")
        
        assert lock_acquired_during_operation, "Lock should be held during clipboard operation"
        assert result == "copied text"

    async def test_clipboard_operations_are_serialized(self):
        """Test that concurrent clipboard operations are serialized by the lock."""
        import main
        main.CLIPBOARD_LOCK = None
        
        execution_order = []
        
        # Create mock tabs
        mock_tab1 = AsyncMock()
        mock_tab2 = AsyncMock()
        
        async def slow_execute_script_1(*args, **kwargs):
            execution_order.append("tab1_start")
            await asyncio.sleep(0.1)  # Simulate some work
            execution_order.append("tab1_end")
            return {"result": {"result": {"value": "text1"}}}
        
        async def slow_execute_script_2(*args, **kwargs):
            execution_order.append("tab2_start")
            await asyncio.sleep(0.1)  # Simulate some work
            execution_order.append("tab2_end")
            return {"result": {"result": {"value": "text2"}}}
        
        mock_tab1.execute_script = slow_execute_script_1
        mock_tab2.execute_script = slow_execute_script_2
        
        mock_button1 = AsyncMock()
        mock_button2 = AsyncMock()
        mock_button1.click = AsyncMock()
        mock_button2.click = AsyncMock()
        
        # Run both operations concurrently
        await asyncio.gather(
            clipboard_click_copy_and_read(mock_tab1, mock_button1, "file1.txt"),
            clipboard_click_copy_and_read(mock_tab2, mock_button2, "file2.txt"),
        )
        
        # Verify operations were serialized (one completes before the other starts)
        # The order should be: tab1_start, tab1_end, tab2_start, tab2_end
        # OR: tab2_start, tab2_end, tab1_start, tab1_end
        # NOT interleaved like: tab1_start, tab2_start, tab1_end, tab2_end
        
        assert len(execution_order) == 4
        
        # Check that operations don't interleave
        if execution_order[0] == "tab1_start":
            assert execution_order[1] == "tab1_end", "Tab1 should complete before tab2 starts"
            assert execution_order[2] == "tab2_start"
            assert execution_order[3] == "tab2_end"
        else:
            assert execution_order[0] == "tab2_start"
            assert execution_order[1] == "tab2_end", "Tab2 should complete before tab1 starts"
            assert execution_order[2] == "tab1_start"
            assert execution_order[3] == "tab1_end"

    async def test_clipboard_click_copy_handles_malformed_response(self):
        """Test that clipboard_click_copy_and_read handles malformed clipboard responses."""
        import main
        main.CLIPBOARD_LOCK = None
        
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock(return_value={"unexpected": "format"})
        
        mock_copy_button = AsyncMock()
        mock_copy_button.click = AsyncMock()
        
        result = await clipboard_click_copy_and_read(mock_tab, mock_copy_button, "test.txt")
        
        assert result == "", "Should return empty string on malformed response"

    async def test_clipboard_write_and_paste_calls_correct_methods(self):
        """Test that clipboard_write_and_paste calls the correct browser methods."""
        import main
        main.CLIPBOARD_LOCK = None
        
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock()
        mock_tab.keyboard = AsyncMock()
        mock_tab.keyboard.hotkey = AsyncMock()
        
        await clipboard_write_and_paste(mock_tab, "test message", "test.txt")
        
        # Verify execute_script was called with clipboard write
        mock_tab.execute_script.assert_called_once()
        call_args = mock_tab.execute_script.call_args[0][0]
        assert "navigator.clipboard.writeText" in call_args
        assert "test message" in call_args
        
        # Verify hotkey was called for paste
        mock_tab.keyboard.hotkey.assert_called_once()

    async def test_lock_released_on_exception(self):
        """Test that the lock is released even if an exception occurs."""
        import main
        main.CLIPBOARD_LOCK = None
        
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock(side_effect=Exception("Test error"))
        mock_tab.keyboard = AsyncMock()
        mock_tab.keyboard.hotkey = AsyncMock()
        
        lock = get_clipboard_lock()
        
        with pytest.raises(Exception):
            await clipboard_write_and_paste(mock_tab, "test", "test.txt")
        
        # Lock should be released after exception
        assert not lock.locked(), "Lock should be released after exception"


class TestClipboardLockIntegration:
    """Integration tests for clipboard lock with simulated concurrent operations."""

    def setup_method(self):
        """Reset the global lock before each test."""
        import main
        main.CLIPBOARD_LOCK = None

    async def test_no_data_loss_with_concurrent_operations(self):
        """Test that no data is lost when multiple operations run concurrently."""
        import main
        main.CLIPBOARD_LOCK = None
        
        results = []
        
        async def mock_clipboard_operation(tab_id: int, expected_value: str):
            """Simulate a clipboard read operation that returns a specific value."""
            mock_tab = AsyncMock()
            mock_tab.execute_script = AsyncMock(return_value={
                "result": {"result": {"value": expected_value}}
            })
            
            mock_button = AsyncMock()
            mock_button.click = AsyncMock()
            
            result = await clipboard_click_copy_and_read(
                mock_tab, mock_button, f"file{tab_id}.txt"
            )
            results.append((tab_id, result))
            return result
        
        # Run 5 concurrent operations
        await asyncio.gather(
            mock_clipboard_operation(1, "content1"),
            mock_clipboard_operation(2, "content2"),
            mock_clipboard_operation(3, "content3"),
            mock_clipboard_operation(4, "content4"),
            mock_clipboard_operation(5, "content5"),
        )
        
        # Verify all operations completed with correct results
        assert len(results) == 5
        
        # Each operation should have received its expected content
        result_dict = {tab_id: content for tab_id, content in results}
        assert result_dict[1] == "content1"
        assert result_dict[2] == "content2"
        assert result_dict[3] == "content3"
        assert result_dict[4] == "content4"
        assert result_dict[5] == "content5"

    async def test_mixed_read_write_operations_serialized(self):
        """Test that mixed read and write operations are properly serialized."""
        import main
        main.CLIPBOARD_LOCK = None
        
        operation_log = []
        
        async def mock_write_operation(tab_id: int):
            """Simulate a clipboard write operation."""
            mock_tab = AsyncMock()
            
            async def log_execute(*args, **kwargs):
                operation_log.append(f"write_{tab_id}_start")
                await asyncio.sleep(0.05)
                operation_log.append(f"write_{tab_id}_end")
            
            mock_tab.execute_script = log_execute
            mock_tab.keyboard = AsyncMock()
            mock_tab.keyboard.hotkey = AsyncMock()
            
            await clipboard_write_and_paste(mock_tab, f"content{tab_id}", f"file{tab_id}.txt")
        
        async def mock_read_operation(tab_id: int):
            """Simulate a clipboard read operation."""
            mock_tab = AsyncMock()
            
            async def log_execute(*args, **kwargs):
                operation_log.append(f"read_{tab_id}_start")
                await asyncio.sleep(0.05)
                operation_log.append(f"read_{tab_id}_end")
                return {"result": {"result": {"value": f"content{tab_id}"}}}
            
            mock_tab.execute_script = log_execute
            
            mock_button = AsyncMock()
            mock_button.click = AsyncMock()
            
            return await clipboard_click_copy_and_read(mock_tab, mock_button, f"file{tab_id}.txt")
        
        # Run mixed operations concurrently
        await asyncio.gather(
            mock_write_operation(1),
            mock_read_operation(2),
            mock_write_operation(3),
        )
        
        # Verify no interleaving - each operation should complete before next starts
        # Check that no "start" appears between another operation's "start" and "end"
        in_progress = None
        for entry in operation_log:
            if entry.endswith("_start"):
                assert in_progress is None, f"Operation {entry} started while {in_progress} was in progress"
                in_progress = entry.replace("_start", "")
            elif entry.endswith("_end"):
                op_name = entry.replace("_end", "")
                assert in_progress == op_name, f"Operation {op_name} ended but {in_progress} was in progress"
                in_progress = None

    async def test_clipboard_read_returns_correct_value(self):
        """Test that clipboard read operation returns the correct value from the mock."""
        import main
        main.CLIPBOARD_LOCK = None
        
        expected_content = "This is the expected clipboard content with special chars: <tag>value</tag>"
        
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock(return_value={
            "result": {"result": {"value": expected_content}}
        })
        
        mock_button = AsyncMock()
        mock_button.click = AsyncMock()
        
        result = await clipboard_click_copy_and_read(mock_tab, mock_button, "test.txt")
        
        assert result == expected_content

    async def test_clipboard_write_escapes_special_characters(self):
        """Test that clipboard write properly handles special characters in JSON."""
        import main
        main.CLIPBOARD_LOCK = None
        
        # Text with special characters that need JSON escaping
        special_text = 'Text with "quotes" and \n newlines and \t tabs'
        
        mock_tab = AsyncMock()
        mock_tab.execute_script = AsyncMock()
        mock_tab.keyboard = AsyncMock()
        mock_tab.keyboard.hotkey = AsyncMock()
        
        await clipboard_write_and_paste(mock_tab, special_text, "test.txt")
        
        # Verify the script was called with properly escaped JSON
        call_args = mock_tab.execute_script.call_args[0][0]
        assert "navigator.clipboard.writeText" in call_args
        # The text should be JSON-encoded (quotes escaped, etc.)
        assert '\\"quotes\\"' in call_args or '"quotes"' not in call_args.replace('\\"', '')


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
