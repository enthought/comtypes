import contextlib
import ctypes
import threading
import unittest as ut
from ctypes import POINTER, WinDLL, c_void_p
from ctypes import c_size_t as SIZE_T
from ctypes.wintypes import BOOL, DWORD, HANDLE, LPCWSTR

from comtypes.logutil import (
    _OutputDebugStringW as OutputDebugStringW,
)
from comtypes.logutil import deprecated


class Test_deprecated(ut.TestCase):
    def test_warning_is_raised(self):
        reason_text = "This is deprecated."

        @deprecated(reason_text)
        def test_func():
            return "success"

        with self.assertWarns(DeprecationWarning) as cm:
            result = test_func()
        self.assertEqual(result, "success")
        self.assertEqual(reason_text, str(cm.warning))


_kernel32 = WinDLL("kernel32")

# https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-createeventw
_CreateEventW = _kernel32.CreateEventW
_CreateEventW.argtypes = [c_void_p, BOOL, BOOL, LPCWSTR]
_CreateEventW.restype = HANDLE

# https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-setevent
_SetEvent = _kernel32.SetEvent
_SetEvent.argtypes = [HANDLE]
_SetEvent.restype = BOOL

# https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitforsingleobject
_WaitForSingleObject = _kernel32.WaitForSingleObject
_WaitForSingleObject.argtypes = [HANDLE, DWORD]
_WaitForSingleObject.restype = DWORD

# https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-createfilemappingw
_CreateFileMappingW = _kernel32.CreateFileMappingW
_CreateFileMappingW.argtypes = [HANDLE, c_void_p, DWORD, DWORD, DWORD, LPCWSTR]
_CreateFileMappingW.restype = HANDLE

# https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-mapviewoffile
_MapViewOfFile = _kernel32.MapViewOfFile
_MapViewOfFile.argtypes = [HANDLE, DWORD, DWORD, DWORD, SIZE_T]
_MapViewOfFile.restype = c_void_p

# https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-unmapviewoffile
_UnmapViewOfFile = _kernel32.UnmapViewOfFile
_UnmapViewOfFile.argtypes = [c_void_p]
_UnmapViewOfFile.restype = BOOL

# https://learn.microsoft.com/en-us/windows/win32/api/handleapi/nf-handleapi-closehandle
_CloseHandle = _kernel32.CloseHandle
_CloseHandle.argtypes = [HANDLE]
_CloseHandle.restype = BOOL

# https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getcurrentprocessid
_GetCurrentProcessId = _kernel32.GetCurrentProcessId
_GetCurrentProcessId.argtypes = []
_GetCurrentProcessId.restype = DWORD


@contextlib.contextmanager
def create_file_mapping(hfile, security, flprotect, size_high, size_low, name):
    """Context manager to creates a Windows file mapping object."""
    handle = _CreateFileMappingW(hfile, security, flprotect, size_high, size_low, name)
    try:
        yield handle
    finally:
        _CloseHandle(handle)


@contextlib.contextmanager
def map_view_of_file(handle, access, offset_high, offset_low, size):
    """Context manager to map a view of a file mapping into the process's
    address space.
    """
    p_view = _MapViewOfFile(handle, access, offset_high, offset_low, size)
    try:
        yield p_view
    finally:
        _UnmapViewOfFile(p_view)


@contextlib.contextmanager
def create_event(security, manual, init, name):
    """Context manager to creates a Windows event object."""
    handle = _CreateEventW(security, manual, init, name)
    try:
        yield handle
    finally:
        _CloseHandle(handle)


@contextlib.contextmanager
def capture_debug_strings(ready, *, interval):
    """Context manager to capture debug strings emitted via `OutputDebugString`.
    Spawns a listener thread to monitor the debug channels.
    """
    captured = []
    finished = threading.Event()
    pid = _GetCurrentProcessId()

    def _listener() -> None:
        # Create/open named events and file mapping for interprocess communication.
        # These objects are part of the Windows Debugging API contract.
        with (
            # "DBWIN_BUFFER_READY": An event signaled by the listener to indicate
            # it's ready to receive debug output. `OutputDebugString` waits for this.
            create_event(None, False, False, "DBWIN_BUFFER_READY") as h_buffer_ready,
            # "DBWIN_DATA_READY": An event signaled by `OutputDebugString` to
            # indicate new data is written to the shared buffer. Listener waits.
            create_event(None, False, False, "DBWIN_DATA_READY") as h_data_ready,
            # "DBWIN_BUFFER": A shared memory region where `OutputDebugString`
            # writes the debug string data.
            create_file_mapping(-1, None, 0x04, 0, 4096, "DBWIN_BUFFER") as h_mapping,
            # Map the shared memory region into the listener's address space
            # for reading the debug strings.
            map_view_of_file(h_mapping, 0x04, 0, 0, 4096) as p_view,
        ):
            ready.set()  # Signal to the main thread that listener is ready.
            # Loop until the main thread signals to finish.
            while not finished.is_set():
                _SetEvent(h_buffer_ready)  # Signal readiness to `OutputDebugString`.
                # Wait for `OutputDebugString` to signal that data is ready.
                if _WaitForSingleObject(h_data_ready, interval) == 0x00000000:
                    # Debug string buffer format: [4 bytes: PID][N bytes: string].
                    # Check if the process ID in the buffer matches the current PID.
                    if ctypes.cast(p_view, POINTER(DWORD)).contents.value == pid:
                        # Extract the null-terminated string, skipping the PID,
                        # and put it into the queue.
                        captured.append(ctypes.string_at(p_view + 4).strip(b"\x00"))

    th = threading.Thread(target=_listener, daemon=True)
    th.start()
    try:
        yield captured
    finally:
        finished.set()
        th.join()


class Test_OutputDebugStringW(ut.TestCase):
    def test(self):
        ready = threading.Event()
        with capture_debug_strings(ready, interval=100) as cap:
            ready.wait(timeout=5)  # Wait for the listener to be ready
            OutputDebugStringW("hello world")
            OutputDebugStringW("test message")
        self.assertEqual(cap[0], b"hello world")
        self.assertEqual(cap[1], b"test message")
