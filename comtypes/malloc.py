from ctypes import HRESULT, POINTER, OleDLL, WinDLL, byref, c_int, c_ulong, c_void_p
from ctypes import c_size_t as SIZE_T
from ctypes.wintypes import BOOL, DWORD, LPVOID
from typing import TYPE_CHECKING, Any

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
    if TYPE_CHECKING:

        def Alloc(self, cb: int) -> int | None: ...
        def Realloc(self, pv: Any, cb: int) -> int | None: ...
        def Free(self, py: Any) -> None: ...
        def GetSize(self, pv: Any) -> int: ...
        def DidAlloc(self, pv: Any) -> int: ...
        def HeapMinimize(self) -> None: ...


class IMallocSpy(IUnknown):
    _iid_ = GUID("{0000001D-0000-0000-C000-000000000046}")
    _methods_ = [
        COMMETHOD([], SIZE_T, "PreAlloc", (["in"], SIZE_T, "cbRequest")),
        COMMETHOD([], c_void_p, "PostAlloc", (["in"], c_void_p, "pActual")),
        COMMETHOD(
            [],
            c_void_p,
            "PreFree",
            (["in"], c_void_p, "pRequest"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD([], None, "PostFree", (["in"], BOOL, "fSpyed")),
        COMMETHOD(
            [],
            SIZE_T,
            "PreRealloc",
            (["in"], c_void_p, "pRequest"),
            (["in"], SIZE_T, "cbRequest"),
            (["out"], POINTER(c_void_p), "ppNewRequest"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD(
            [],
            c_void_p,
            "PostRealloc",
            (["in"], c_void_p, "pActual"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD(
            [],
            c_void_p,
            "PreGetSize",
            (["in"], c_void_p, "pRequest"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD(
            [],
            SIZE_T,
            "PostGetSize",
            (["in"], SIZE_T, "cbActual"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD(
            [],
            c_void_p,
            "PreDidAlloc",
            (["in"], c_void_p, "pRequest"),
            (["in"], BOOL, "fSpyed"),
        ),
        COMMETHOD(
            [],
            c_int,
            "PostDidAlloc",
            (["in"], c_void_p, "pRequest"),
            (["in"], BOOL, "fSpyed"),
            (["in"], c_int, "fActual"),
        ),
        COMMETHOD([], None, "PreHeapMinimize"),
        COMMETHOD([], None, "PostHeapMinimize"),
    ]


_ole32 = OleDLL("ole32")

_CoGetMalloc = _ole32.CoGetMalloc
_CoGetMalloc.argtypes = [DWORD, POINTER(POINTER(IMalloc))]
_CoGetMalloc.restype = HRESULT

_CoRegisterMallocSpy = _ole32.CoRegisterMallocSpy
_CoRegisterMallocSpy.argtypes = [POINTER(IMallocSpy)]
_CoRegisterMallocSpy.restype = HRESULT

_CoRevokeMallocSpy = _ole32.CoRevokeMallocSpy
_CoRevokeMallocSpy.argtypes = []
_CoRevokeMallocSpy.restype = HRESULT

_ole32_nohresult = WinDLL("ole32")

_CoTaskMemAlloc = _ole32_nohresult.CoTaskMemAlloc
_CoTaskMemAlloc.argtypes = [SIZE_T]
_CoTaskMemAlloc.restype = LPVOID


def CoGetMalloc(dwMemContext: int = 1) -> IMalloc:
    """Retrieves a pointer to the default OLE task memory allocator.

    https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cogetmalloc
    """
    malloc = POINTER(IMalloc)()
    _CoGetMalloc(
        dwMemContext,  # This parameter must be 1.
        byref(malloc),
    )
    return malloc  # type: ignore
