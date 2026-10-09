# logutil.py
import logging
from ctypes import WinDLL
from ctypes.wintypes import LPCSTR, LPCWSTR

_kernel32 = WinDLL("kernel32")

_OutputDebugStringA = _kernel32.OutputDebugStringA
_OutputDebugStringA.argtypes = [LPCSTR]
_OutputDebugStringA.restype = None

_OutputDebugStringW = _kernel32.OutputDebugStringW
_OutputDebugStringW.argtypes = [LPCWSTR]
_OutputDebugStringW.restype = None


class NTDebugHandler(logging.Handler):
    def emit(
        self,
        record,
        writeW=_OutputDebugStringW,
    ):
        text = self.format(record)
        writeW(text + "\n")


logging.NTDebugHandler = NTDebugHandler
