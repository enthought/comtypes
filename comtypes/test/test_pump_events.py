import gc
import unittest
from ctypes import c_void_p, cast
from unittest.mock import patch

from comtypes.client import PumpEvents, _events


class PumpEventsTest(unittest.TestCase):
    def test_removes_the_registered_callback_before_closing_the_event(self):
        operations = []
        callbacks = []
        close_handle = _events._CloseHandle

        def record_handler(handler, add):
            callbacks.append(handler)
            operations.append((add, cast(handler, c_void_p).value))
            return 1

        def record_close(handle):
            operations.append("close")
            return close_handle(handle)

        with (
            patch.object(_events, "_SetConsoleCtrlHandler", record_handler),
            patch.object(_events, "_CloseHandle", record_close),
        ):
            PumpEvents(0.001)
        address = cast(callbacks[0], c_void_p).value
        self.assertEqual(operations, [(1, address), (0, address), "close"])

    def test_other_hresult_failures_are_reported(self):
        invalid_argument = -2147024809
        with patch.object(
            _events, "_CoWaitForMultipleHandles", return_value=invalid_argument
        ):
            with self.assertRaises(OSError) as caught:
                PumpEvents(0.001)
        self.assertEqual(caught.exception.winerror, invalid_argument)

    def test_signaled_event_raises_keyboard_interrupt(self):
        with patch.object(_events, "_CoWaitForMultipleHandles", return_value=0):
            with self.assertRaises(KeyboardInterrupt):
                PumpEvents(0.001)

    def test_pump_events_doesnt_leak_cycles(self):
        gc.collect()
        for i in range(3):
            PumpEvents(0.05)
            ncycles = gc.collect()
            self.assertEqual(ncycles, 0)


if __name__ == "__main__":
    unittest.main()
