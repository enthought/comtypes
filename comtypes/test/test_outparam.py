import unittest
from ctypes import (
    HRESULT,
    POINTER,
    OleDLL,
    WinDLL,
    byref,
    c_int,
    c_size_t,
    c_ulong,
    c_void_p,
    c_wchar,
    c_wchar_p,
    cast,
    memmove,
    sizeof,
    wstring_at,
)
from ctypes.wintypes import DWORD, LPVOID
from unittest.mock import patch

from comtypes import COMMETHOD, GUID, IUnknown
from comtypes.GUID import _CoTaskMemFree

text_type = str


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

malloc = POINTER(IMalloc)()
_CoGetMalloc(1, byref(malloc))
assert bool(malloc)


def from_outparm(self):
    if not self:
        return None
    result = wstring_at(self)
    if not malloc.DidAlloc(self):
        raise ValueError("memory was NOT allocated by CoTaskMemAlloc")
    _CoTaskMemFree(self)
    return result


def comstring(text, typ=c_wchar_p):
    text = text_type(text)
    size = (len(text) + 1) * sizeof(c_wchar)
    mem = _CoTaskMemAlloc(size)
    print("malloc'd 0x%x, %d bytes" % (mem, size))
    ptr = cast(mem, typ)
    memmove(mem, text, size)
    return ptr


class Test(unittest.TestCase):
    @unittest.skip("This fails for reasons I don't understand yet")
    # TODO untested changes; this was modified because it had global effects on other tests
    @patch.object(c_wchar_p, "__ctypes_from_outparam__", from_outparm)
    def test_c_char(self):
        # ptr = c_wchar_p("abc")
        # self.failUnlessEqual(ptr.__ctypes_from_outparam__(),
        #                         "abc")

        # p = BSTR("foo bar spam")

        x = comstring("Hello, World")
        y = comstring("foo bar")
        z = comstring("spam, spam, and spam")

        # (x.__ctypes_from_outparam__(), x.__ctypes_from_outparam__())
        print((x.__ctypes_from_outparam__(), None))  # x.__ctypes_from_outparam__())

        # print comstring("Hello, World", c_wchar_p).__ctypes_from_outparam__()
        # print comstring("Hello, World", c_wchar_p).__ctypes_from_outparam__()
        # print comstring("Hello, World", c_wchar_p).__ctypes_from_outparam__()
        # print comstring("Hello, World", c_wchar_p).__ctypes_from_outparam__()


if __name__ == "__main__":
    unittest.main()
