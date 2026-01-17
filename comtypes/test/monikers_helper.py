from ctypes import HRESULT, POINTER, OleDLL, c_wchar_p
from ctypes.wintypes import DWORD

from comtypes import IUnknown

MKSYS_ITEMMONIKER = 4

ROTFLAGS_ALLOWANYCLIENT = 1

LPOLESTR = LPCOLESTR = c_wchar_p

_ole32 = OleDLL("ole32")

_CreateItemMoniker = _ole32.CreateItemMoniker
_CreateItemMoniker.argtypes = [LPCOLESTR, LPCOLESTR, POINTER(POINTER(IUnknown))]
_CreateItemMoniker.restype = HRESULT

_CreateBindCtx = _ole32.CreateBindCtx
_CreateBindCtx.argtypes = [DWORD, POINTER(POINTER(IUnknown))]
_CreateBindCtx.restype = HRESULT

_GetRunningObjectTable = _ole32.GetRunningObjectTable
_GetRunningObjectTable.argtypes = [DWORD, POINTER(POINTER(IUnknown))]
_GetRunningObjectTable.restype = HRESULT
