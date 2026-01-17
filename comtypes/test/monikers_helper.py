from ctypes import HRESULT, POINTER, OleDLL, c_wchar_p
from ctypes.wintypes import DWORD

from comtypes import GUID, IUnknown

# https://learn.microsoft.com/en-us/windows/win32/api/objidl/ne-objidl-mksys
MKSYS_ITEMMONIKER = 4

CLSID_AntiMoniker = GUID("{00000305-0000-0000-c000-000000000046}")
CLSID_ItemMoniker = GUID("{00000304-0000-0000-c000-000000000046}")

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

# Common COM Errors from Moniker/Binding Context operations
MK_E_UNAVAILABLE = -2147221021  # 0x800401E3
