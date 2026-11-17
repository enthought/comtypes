import gc
import unittest
import unittest as ut
from ctypes import HRESULT, POINTER, OleDLL, byref
from ctypes.wintypes import DWORD, HANDLE, LPWSTR
from pathlib import Path
from unittest.mock import MagicMock

from comtypes import GUID, COMObject, hresult
from comtypes.malloc import (
    CoGetMalloc,
    IMallocSpy,
    _CoRegisterMallocSpy,
    _CoRevokeMallocSpy,
    _CoTaskMemAlloc,
    _CoTaskMemFree,
)

# Constants
# KNOWNFOLDERID
# https://learn.microsoft.com/en-us/windows/win32/shell/knownfolderid
FOLDERID_System = GUID("{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}")
# https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/ne-shlobj_core-known_folder_flag
KF_FLAG_DEFAULT = 0x00000000

_shell32 = OleDLL("shell32")
_SHGetKnownFolderPath = _shell32.SHGetKnownFolderPath
_SHGetKnownFolderPath.argtypes = [
    POINTER(GUID),  # rfid
    DWORD,  # dwFlags
    HANDLE,  # hToken
    POINTER(LPWSTR),  # ppszPath
]
_SHGetKnownFolderPath.restype = HRESULT


class Test(ut.TestCase):
    def test_Realloc(self):
        malloc = CoGetMalloc()
        size1 = 4
        ptr1 = malloc.Alloc(size1)
        self.assertEqual(malloc.DidAlloc(ptr1), 1)
        self.assertEqual(malloc.GetSize(ptr1), size1)
        size2 = size1 - 1
        ptr2 = malloc.Realloc(ptr1, size2)
        self.assertEqual(malloc.DidAlloc(ptr2), 1)
        self.assertEqual(malloc.GetSize(ptr2), size2)
        size3 = size1 + 1
        ptr3 = malloc.Realloc(ptr2, size3)
        self.assertEqual(malloc.DidAlloc(ptr3), 1)
        self.assertEqual(malloc.GetSize(ptr3), size3)
        malloc.Free(ptr3)
        self.assertEqual(malloc.DidAlloc(ptr3), 0)
        malloc.HeapMinimize()
        del ptr3

    def test_SHGetKnownFolderPath(self):
        ptr = LPWSTR()
        hr = _SHGetKnownFolderPath(
            byref(FOLDERID_System), KF_FLAG_DEFAULT, None, byref(ptr)
        )
        self.assertEqual(hr, hresult.S_OK)
        self.assertIsInstance(ptr.value, str)
        self.assertTrue(Path(ptr.value).exists())  # type: ignore
        malloc = CoGetMalloc()
        self.assertEqual(malloc.DidAlloc(ptr), 1)
        self.assertGreater(malloc.GetSize(ptr), 0)
        _CoTaskMemFree(ptr)
        self.assertEqual(malloc.DidAlloc(ptr), 0)
        malloc.HeapMinimize()
        del ptr


class MockMallocSpy(COMObject):
    """IMallocSpy implementation that records every hook call."""

    _com_interfaces_ = [IMallocSpy]

    def __init__(self):
        super().__init__()
        self.mock = MagicMock()
        # Each hook forwards to a child of `self.mock`, so `mock.method_calls`
        # holds the hooks in call order, while each child keeps its own
        # arguments.
        self.PreAlloc: MagicMock = self.mock.PreAlloc
        self.PostAlloc: MagicMock = self.mock.PostAlloc
        self.PreFree: MagicMock = self.mock.PreFree
        self.PostFree: MagicMock = self.mock.PostFree
        self.PreRealloc: MagicMock = self.mock.PreRealloc
        self.PostRealloc: MagicMock = self.mock.PostRealloc
        self.PreGetSize: MagicMock = self.mock.PreGetSize
        self.PostGetSize: MagicMock = self.mock.PostGetSize
        self.PreDidAlloc: MagicMock = self.mock.PreDidAlloc
        self.PostDidAlloc: MagicMock = self.mock.PostDidAlloc
        self.PreHeapMinimize: MagicMock = self.mock.PreHeapMinimize
        self.PostHeapMinimize: MagicMock = self.mock.PostHeapMinimize

    @property
    def call_names(self) -> list[str]:
        """Return the names of the called hooks, in call order."""
        return [name for name, _, _ in self.mock.method_calls]

    def IMallocSpy_PreAlloc(self, this, cbRequest):
        self.PreAlloc(this, cbRequest)
        return cbRequest

    def IMallocSpy_PostAlloc(self, this, pActual):
        self.PostAlloc(this, pActual)
        return pActual

    def IMallocSpy_PreFree(self, this, pRequest, fSpyed):
        self.PreFree(this, pRequest, fSpyed)
        return pRequest

    def IMallocSpy_PostFree(self, this, fSpyed):
        self.PostFree(this, fSpyed)

    def IMallocSpy_PreRealloc(self, this, pRequest, cbRequest, ppNewRequest, fSpyed):
        self.PreRealloc(this, pRequest, cbRequest, ppNewRequest, fSpyed)
        if ppNewRequest:
            ppNewRequest[0] = pRequest
        return cbRequest

    def IMallocSpy_PostRealloc(self, this, pActual, fSpyed):
        self.PostRealloc(this, pActual, fSpyed)
        return pActual

    def IMallocSpy_PreGetSize(self, this, pRequest, fSpyed):
        self.PreGetSize(this, pRequest, fSpyed)
        return pRequest

    def IMallocSpy_PostGetSize(self, this, cbActual, fSpyed):
        self.PostGetSize(this, cbActual, fSpyed)
        return cbActual

    def IMallocSpy_PreDidAlloc(self, this, pRequest, fSpyed):
        self.PreDidAlloc(this, pRequest, fSpyed)
        return pRequest

    def IMallocSpy_PostDidAlloc(self, this, pRequest, fSpyed, fActual):
        self.PostDidAlloc(this, pRequest, fSpyed, fActual)
        return fActual

    def IMallocSpy_PreHeapMinimize(self, this):
        self.PreHeapMinimize(this)

    def IMallocSpy_PostHeapMinimize(self, this):
        self.PostHeapMinimize(this)


class Test_IMallocSpy(unittest.TestCase):
    def create_spy(self) -> MockMallocSpy:
        self.addCleanup(_CoRevokeMallocSpy)
        spy = MockMallocSpy()
        _CoRegisterMallocSpy(spy)
        return spy

    def test_task_mem_alloc_free_order(self):
        spy = self.create_spy()
        ptr = _CoTaskMemAlloc(64)
        _CoTaskMemFree(ptr)
        self.assertEqual(
            spy.call_names, ["PreAlloc", "PostAlloc", "PreFree", "PostFree"]
        )

    def test_imalloc_alloc_free_order(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(64)
        malloc.Free(ptr)
        self.assertEqual(
            spy.call_names, ["PreAlloc", "PostAlloc", "PreFree", "PostFree"]
        )

    def test_realloc_order(self):
        malloc = CoGetMalloc()
        ptr = malloc.Alloc(32)
        spy = self.create_spy()
        new_ptr = malloc.Realloc(ptr, 256)
        self.assertEqual(spy.call_names, ["PreRealloc", "PostRealloc"])
        malloc.Free(new_ptr)

    def test_getsize_order(self):
        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        spy = self.create_spy()
        malloc.GetSize(ptr)
        self.assertEqual(spy.call_names, ["PreGetSize", "PostGetSize"])
        malloc.Free(ptr)

    def test_didalloc_order(self):
        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        spy = self.create_spy()
        malloc.DidAlloc(ptr)
        self.assertEqual(spy.call_names, ["PreDidAlloc", "PostDidAlloc"])
        malloc.Free(ptr)

    def test_heapminimize_order(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        malloc.HeapMinimize()
        self.assertEqual(spy.call_names, ["PreHeapMinimize", "PostHeapMinimize"])

    def test_non_com_activity_fires_nothing(self):
        spy = self.create_spy()
        data = [i for i in range(1000)]
        del data
        gc.collect()
        self.assertEqual(spy.call_names, [])

    def test_prealloc_receives_size(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(12345)
        malloc.Free(ptr)
        spy.PreAlloc.assert_called_once()
        # args: (this, cbRequest);
        # `this` is the COM interface pointer (opaque `int`).
        # It is constant for the spy instance; not relevant to the test.
        _, cbRequest = spy.PreAlloc.call_args.args
        self.assertEqual(cbRequest, 12345)

    def test_postalloc_receives_pointer(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.Free(ptr)
        spy.PostAlloc.assert_called_once()
        _, pActual = spy.PostAlloc.call_args.args
        # `pActual` and `ptr` are Python `int`s (not `None` here).
        # `ctypes` converts non-NULL `c_void_p` to `int`; NULL becomes `None`.
        # They hold the address value; comparing ints is safe after `Free`.
        self.assertEqual(pActual, ptr)

    def test_prefree_receives_pointer(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.Free(ptr)
        spy.PreFree.assert_called_once()
        _, pRequest, fSpyed = spy.PreFree.call_args.args
        self.assertEqual(pRequest, ptr)
        self.assertEqual(fSpyed, 1)

    def test_postfree_receives_fspyed(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.Free(ptr)
        spy.PostFree.assert_called_once()
        _, fSpyed = spy.PostFree.call_args.args
        self.assertEqual(fSpyed, 1)

    def test_prerealloc_receives_old_pointer_and_new_size(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        new_ptr = malloc.Realloc(ptr, 200)
        malloc.Free(new_ptr)
        spy.PreRealloc.assert_called_once()
        # args: (this, pRequest, cbRequest, ppNewRequest, fSpyed);
        # `ppNewRequest` is an out-param (`POINTER(c_void_p)`).
        # Reading it gives an uninitialized `ctypes` pointer;
        # dereferencing crashes.
        _, pRequest, cbRequest, _, fSpyed = spy.PreRealloc.call_args.args
        self.assertEqual(pRequest, ptr)
        self.assertEqual(cbRequest, 200)
        self.assertEqual(fSpyed, 1)

    def test_postrealloc_receives_pointer(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        new_ptr = malloc.Realloc(ptr, 200)
        malloc.Free(new_ptr)
        spy.PostRealloc.assert_called_once()
        _, pActual, fSpyed = spy.PostRealloc.call_args.args
        self.assertEqual(pActual, new_ptr)
        self.assertEqual(fSpyed, 1)

    def test_pregetsize_receives_pointer(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.GetSize(ptr)
        malloc.Free(ptr)
        spy.PreGetSize.assert_called_once()
        _, pRequest, fSpyed = spy.PreGetSize.call_args.args
        self.assertEqual(pRequest, ptr)
        self.assertEqual(fSpyed, 1)

    def test_postgetsize_receives_size(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.GetSize(ptr)
        malloc.Free(ptr)
        spy.PostGetSize.assert_called_once()
        _, cbActual, fSpyed = spy.PostGetSize.call_args.args
        self.assertEqual(cbActual, 100)
        self.assertEqual(fSpyed, 1)

    def test_predidalloc_receives_pointer(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.DidAlloc(ptr)
        malloc.Free(ptr)
        spy.PreDidAlloc.assert_called_once()
        _, pRequest, fSpyed = spy.PreDidAlloc.call_args.args
        self.assertEqual(pRequest, ptr)
        self.assertEqual(fSpyed, 1)

    def test_postdidalloc_receives_result(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        ptr = malloc.Alloc(100)
        malloc.DidAlloc(ptr)
        malloc.Free(ptr)
        spy.PostDidAlloc.assert_called_once()
        _, pRequest, fSpyed, fActual = spy.PostDidAlloc.call_args.args
        self.assertEqual(pRequest, ptr)
        self.assertEqual(fSpyed, 1)
        self.assertEqual(fActual, 1)

    def test_preheapminimize_called(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        malloc.HeapMinimize()
        spy.PreHeapMinimize.assert_called_once()

    def test_postheapminimize_called(self):
        malloc = CoGetMalloc()
        spy = self.create_spy()
        malloc.HeapMinimize()
        spy.PostHeapMinimize.assert_called_once()

    def test_register_twice_raises_objisreg(self):
        spy1 = MockMallocSpy()
        hr = _CoRegisterMallocSpy(spy1)
        self.assertEqual(hr, hresult.S_OK)
        spy2 = MockMallocSpy()
        with self.assertRaises(OSError) as cm:
            _CoRegisterMallocSpy(spy2)
        self.assertEqual(cm.exception.winerror, hresult.CO_E_OBJISREG)
        _CoRevokeMallocSpy()

    def test_revoke_without_register_raises_objnotreg(self):
        with self.assertRaises(OSError) as cm:
            _CoRevokeMallocSpy()
        self.assertEqual(cm.exception.winerror, hresult.CO_E_OBJNOTREG)

    def test_revoke_twice_raises_objnotreg(self):
        spy = MockMallocSpy()
        _CoRegisterMallocSpy(spy)
        hr = _CoRevokeMallocSpy()
        self.assertEqual(hr, hresult.S_OK)
        with self.assertRaises(OSError) as cm:
            _CoRevokeMallocSpy()
        self.assertEqual(cm.exception.winerror, hresult.CO_E_OBJNOTREG)
