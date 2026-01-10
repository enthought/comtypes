from ctypes import HRESULT, POINTER, OleDLL, WinDLL, c_int, c_size_t, c_ulong, c_void_p
from ctypes.wintypes import DWORD, LPVOID

from comtypes import COMMETHOD, GUID, IUnknown
from comtypes.GUID import _CoTaskMemFree as _CoTaskMemFree


class IMalloc(IUnknown):
    _iid_ = GUID("{00000002-0000-0000-C000-000000000046}")
    _methods_ = [
        COMMETHOD([], c_void_p, "Alloc", ([], c_ulong, "cb")),
        COMMETHOD([], c_void_p, "Realloc", ([], c_void_p, "pv"), ([], c_ulong, "cb")),
        COMMETHOD([], None, "Free", ([], c_void_p, "py")),
        COMMETHOD([], c_ulong, "GetSize", ([], c_void_p, "pv")),
        COMMETHOD([], c_int, "DidAlloc", ([], c_void_p, "pv")),
        COMMETHOD([], None, "HeapMinimize"),  # 25
    ]


_ole32 = OleDLL("ole32")

_CoGetMalloc = _ole32.CoGetMalloc
_CoGetMalloc.argtypes = [DWORD, POINTER(POINTER(IMalloc))]
_CoGetMalloc.restype = HRESULT

_ole32_nohresult = WinDLL("ole32")

SIZE_T = c_size_t
_CoTaskMemAlloc = _ole32_nohresult.CoTaskMemAlloc
_CoTaskMemAlloc.argtypes = [SIZE_T]
_CoTaskMemAlloc.restype = LPVOID
