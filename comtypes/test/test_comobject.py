import ctypes
import unittest as ut
from _ctypes import COMError
from ctypes import POINTER, FormatError, WinError, byref, pointer
from typing import TYPE_CHECKING
from unittest import mock

import comtypes.client
from comtypes import (
    CLSCTX_SERVER,
    COMMETHOD,
    GUID,
    STDMETHOD,
    COMObject,
    IPersist,
    IUnknown,
    hresult,
)
from comtypes._post_coinit.misc import _CoCreateInstance
from comtypes.automation import IDispatch
from comtypes.errorinfo import ReportError
from comtypes.server import IClassFactory
from comtypes.typeinfo import GUIDKIND_DEFAULT_SOURCE_DISP_IID

if TYPE_CHECKING:
    from comtypes import hints  # type: ignore

comtypes.client.GetModule("UIAutomationCore.dll")
comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting as scrrun
from comtypes.gen import UIAutomationClient as uiac
from comtypes.gen import stdole


class Test_QueryInterface(ut.TestCase):
    def test_custom_interface(self):
        iuia = uiac.CUIAutomation().QueryInterface(uiac.IUIAutomation)
        self.assertTrue(bool(iuia))
        self.assertEqual(iuia.AddRef(), 2)
        punk = iuia.QueryInterface(IUnknown)
        self.assertEqual(iuia.AddRef(), 4)
        self.assertEqual(punk.AddRef(), 5)
        self.assertEqual(iuia.Release(), 4)
        del iuia
        self.assertEqual(punk.Release(), 2)
        self.assertEqual(punk.Release(), 1)

    def test_dispatch_interface(self):
        dic = scrrun.Dictionary().QueryInterface(scrrun.IDictionary)
        self.assertTrue(bool(dic))
        self.assertEqual(dic.GetTypeInfoCount(), 1)
        self.assertEqual(
            dic.GetTypeInfo(0).GetTypeAttr().guid,
            scrrun.IDictionary._iid_,
        )
        self.assertEqual(dic.GetIDsOfNames("Add"), [1])


class Test_IUnknown_QueryInterface(ut.TestCase):
    def test_e_pointer(self):
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(scrrun.IDictionary._iid_), ctypes.c_void_p()
        )
        self.assertEqual(hr, hresult.E_POINTER)

    def test_e_no_interface(self):
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(uiac.IUIAutomation._iid_), ctypes.c_void_p()
        )
        self.assertEqual(hr, hresult.E_NOINTERFACE)

    def test_valid_pointer(self):
        ptr = ctypes.c_void_p()
        _CoCreateInstance(
            byref(scrrun.Dictionary._reg_clsid_),
            None,
            CLSCTX_SERVER,
            byref(scrrun.IDictionary._iid_),
            byref(ptr),
        )
        before_val = ptr.value
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(scrrun.IDictionary._iid_), ptr
        )
        self.assertEqual(hr, hresult.S_OK)
        self.assertEqual(ptr.value, before_val)

    def test_valid_interface(self):
        dic = POINTER(IDispatch)()
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(scrrun.IDictionary._iid_), byref(dic)
        )
        self.assertEqual(hr, hresult.S_OK)
        self.assertEqual(dic.AddRef(), 2)  # type: ignore
        self.assertEqual(dic.Release(), 1)  # type: ignore
        self.assertEqual(dic.GetTypeInfoCount(), 1)  # type: ignore


class Test_IUnknown_AddRef_IUnknown_Release(ut.TestCase):
    def test(self):
        cuia = uiac.CUIAutomation()
        self.assertEqual(cuia.IUnknown_AddRef(None), 1)
        self.assertEqual(cuia.IUnknown_AddRef(None), 2)
        with mock.patch.object(cuia, "_final_release_") as release:
            self.assertEqual(cuia.IUnknown_Release(None), 1)
            release.assert_not_called()
            self.assertEqual(cuia.IUnknown_Release(None), 0)
            release.assert_called_once_with()


class Test_ISupportErrorInfo_InterfaceSupportsErrorInfo(ut.TestCase):
    def test_s_ok(self):
        cuia = uiac.CUIAutomation()
        self.assertEqual(
            cuia.ISupportErrorInfo_InterfaceSupportsErrorInfo(
                None, pointer(IUnknown._iid_)
            ),
            hresult.S_OK,
        )
        self.assertEqual(
            cuia.ISupportErrorInfo_InterfaceSupportsErrorInfo(
                None, pointer(uiac.IUIAutomation._iid_)
            ),
            hresult.S_OK,
        )

    def test_s_false(self):
        cuia = uiac.CUIAutomation()
        self.assertEqual(
            cuia.ISupportErrorInfo_InterfaceSupportsErrorInfo(
                None, pointer(IDispatch._iid_)
            ),
            hresult.S_FALSE,
        )
        self.assertEqual(
            cuia.ISupportErrorInfo_InterfaceSupportsErrorInfo(
                None, pointer(scrrun.IDictionary._iid_)
            ),
            hresult.S_FALSE,
        )


class Test_IProvideClassInfo_GetClassInfo(ut.TestCase):
    def test(self):
        tinfo = uiac.CUIAutomation().IProvideClassInfo_GetClassInfo()
        self.assertEqual(tinfo.GetTypeAttr().guid, uiac.CUIAutomation._reg_clsid_)


class Test_IProvideClassInfo2_GetGUID(ut.TestCase):
    def test(self):
        obj = stdole.StdFont()
        actual = obj.IProvideClassInfo2_GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID)
        self.assertEqual(actual, stdole.FontEvents._iid_)


class Test_IPersist_GetClassID(ut.TestCase):
    def test(self):
        self.assertEqual(
            uiac.CUIAutomation().IPersist_GetClassID(),
            uiac.CUIAutomation._reg_clsid_,
        )


class Test_Vtbl(ut.TestCase):
    def test_com_pointers(self):
        stdpic = stdole.StdPicture()
        cuia = uiac.CUIAutomation()
        self.assertIs(
            type(stdpic._com_pointers_[IPersist._iid_]),
            type(cuia._com_pointers_[IPersist._iid_]),
        )
        self.assertIn(IDispatch._iid_, stdpic._com_pointers_)
        self.assertNotIn(IDispatch._iid_, cuia._com_pointers_)
        self.assertIn(stdole.IPictureDisp._iid_, stdpic._com_pointers_)
        self.assertIn(uiac.IUIAutomation._iid_, cuia._com_pointers_)


class Test_CustomImplementation(ut.TestCase):
    def test_raises_comerror(self):
        ERR_DESC = "Simulated COMError"
        ERR_HELPFILE = "test.hlp"
        ERR_HELPCTX = 42

        class MyClassFactory(COMObject):
            _com_interfaces_ = [IClassFactory]
            _reg_clsid_ = GUID.create_new()

            def CreateInstance(self, this, punkOuter, riid, ppv):
                return ReportError(
                    f"{ERR_DESC}: CreateInstance",
                    IClassFactory._iid_,
                    self._reg_clsid_,
                    ERR_HELPFILE,
                    ERR_HELPCTX,
                    hresult.E_UNEXPECTED,
                )

            def LockServer(self, this, fLock):
                return ReportError(
                    f"{ERR_DESC}: LockServer",
                    IClassFactory._iid_,
                    self._reg_clsid_,
                    ERR_HELPFILE,
                    ERR_HELPCTX,
                    hresult.E_FAIL,
                )

        # Get a COM pointer to the interface of our custom object.
        cf = MyClassFactory().QueryInterface(IClassFactory)
        # calling `LockServer`
        with self.assertRaises(COMError) as cm:
            cf.CreateInstance(interface=IUnknown)
        self.assertEqual(cm.exception.hresult, hresult.E_UNEXPECTED)
        self.assertEqual(cm.exception.text, FormatError(hresult.E_UNEXPECTED))
        self.assertEqual(
            cm.exception.details,
            (f"{ERR_DESC}: CreateInstance", None, ERR_HELPFILE, ERR_HELPCTX, None),
        )
        # calling `LockServer`
        with self.assertRaises(COMError) as cm:
            cf.LockServer(True)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(cm.exception.text, FormatError(hresult.E_FAIL))
        self.assertEqual(
            cm.exception.details,
            (f"{ERR_DESC}: LockServer", None, ERR_HELPFILE, ERR_HELPCTX, None),
        )


class Test_HRESULT_via_COMObject(ut.TestCase):
    class ITestHresultFromWin32(IUnknown):
        # A minimal interface with one `STDMETHOD` method and one `COMMETHOD`
        # method.  `STDMETHOD` sets `paramflags=None`, so its methods always go
        # through `catch_errors`; `COMMETHOD` sets `paramflags`, so a method
        # implementation without a `this` parameter goes through `hack`'s
        # `call_without_this`.  This lets the tests reach both conversion paths
        # without depending on a real COM interface.
        _iid_ = GUID.create_new()
        _methods_ = [
            STDMETHOD(ctypes.HRESULT, "StdMethod", [ctypes.c_int]),
            COMMETHOD([], ctypes.HRESULT, "ComMethod", (["in"], ctypes.c_int, "arg")),
        ]
        if TYPE_CHECKING:

            def StdMethod(self, arg) -> hints.Hresult: ...
            def ComMethod(self, arg) -> hints.Hresult: ...

    def test_oserror_is_converted(self):
        # A COM method implementation that raises `OSError` is caught by
        # `catch_errors`, which converts the Windows error code to an HRESULT
        # with `HRESULT_FROM_WIN32`.  `ERROR_INVALID_PARAMETER` (87) must map to
        # `E_INVALIDARG` (0x80070057), so the COM caller observes that HRESULT.
        #
        # On Windows, `.winerror` must be populated for `_vtbl.winerror()` to
        # extract the Win32 error code; otherwise it falls back to `E_FAIL`.
        # Two idiomatic ways to raise with `.winerror` set:
        #   - `ctypes.WinError(code)`
        #     - fills `.strerror` via `FormatMessage`
        #   - `OSError(0, strerror, None, code)`
        #     - 4-arg form; `strerror` is caller-supplied
        # The 2-argument form OSError(errno, strerror) leaves `.winerror = None`
        # and puts the value in `.errno` instead.
        ERROR_INVALID_PARAMETER = 87

        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                raise WinError(ERROR_INVALID_PARAMETER)

            def ComMethod(self, arg):
                return hresult.S_OK

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="ERROR") as logs:
            with self.assertRaises(COMError) as cm:
                obj.StdMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.StdMethod implementation:",
        )
        self.assertEqual(obj.ComMethod(0), hresult.S_OK)

    def test_oserror_in_call_without_this_is_converted(self):
        # A COM method that does not take a `this` parameter is wrapped by
        # `hack`'s `call_without_this`, whose `except OSError` handler converts
        # the Windows error code with `HRESULT_FROM_WIN32`.
        # `ERROR_INVALID_PARAMETER` (87) must map to `E_INVALIDARG`
        # (0x80070057).
        ERROR_INVALID_PARAMETER = 87

        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                return hresult.S_OK

            def ComMethod(self, arg):
                raise WinError(ERROR_INVALID_PARAMETER)

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="ERROR") as logs:
            with self.assertRaises(COMError) as cm:
                obj.ComMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.ComMethod implementation:",
        )

    def test_zero_winerror_is_reported_as_failure(self):
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                raise OSError(0, "simulated", None, 0)

            def ComMethod(self, arg):
                return hresult.S_OK

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="WARNING") as logs:
            with self.assertRaises(COMError) as cm:
                obj.StdMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.StdMethod implementation:",
        )
        self.assertEqual(
            logs.records[1].getMessage(),
            "HRESULT_FROM_WIN32(0) is S_OK; using E_FAIL instead so "
            "that the failure is not reported as success",
        )

    def test_zero_winerror_in_call_without_this_is_reported_as_failure(self):
        # A `0` winerror must NOT be reported as success.
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                return hresult.S_OK

            def ComMethod(self, arg):
                raise OSError(0, "simulated", None, 0)

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="WARNING") as logs:
            with self.assertRaises(COMError) as cm:
                obj.ComMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.ComMethod implementation:",
        )
        self.assertEqual(
            logs.records[1].getMessage(),
            "HRESULT_FROM_WIN32(0) is S_OK; using E_FAIL instead so "
            "that the failure is not reported as success",
        )

    def test_comerror_hresult(self):
        # `HRESULT_FROM_WIN32(0)` returns `S_OK` (0) to match the Windows macro.
        # However, `catch_errors` runs on an error path. There, a `0` code
        # (like `COMError(0, ...)`) looks like success and can hide the error.
        # To prevent this, the error handler uses `E_FAIL` instead and warns
        # the caller. This is the same fallback `winerror` uses when a
        # `WindowsError` has no code.
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                raise COMError(0, "simulated", (None, None, None, 0, None))

            def ComMethod(self, arg):
                return hresult.S_OK

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="WARNING") as logs:
            with self.assertRaises(COMError) as cm:
                obj.StdMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.StdMethod implementation:",
        )
        self.assertEqual(
            logs.records[1].getMessage(),
            "HRESULT_FROM_WIN32(0) is S_OK; using E_FAIL instead so "
            "that the failure is not reported as success",
        )

    def test_comerror_hresult_in_call_without_this(self):
        # A COM method that does not take a `this` parameter is wrapped by
        # `hack`'s `call_without_this`, whose `except COMError` handler converts
        # the hresult with `HRESULT_FROM_WIN32`.  A negative hresult (a normal
        # COM error) is already in HRESULT form, so it is returned unchanged.
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                return hresult.S_OK

            def ComMethod(self, arg):
                raise COMError(hresult.E_FAIL, "simulated", (None, None, None, 0, None))

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="ERROR") as logs:
            with self.assertRaises(COMError) as cm:
                obj.ComMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.ComMethod implementation:",
        )

    def test_comerror_zero_hresult_in_call_without_this(self):
        # A COMError with `0` hresult must NOT be reported as success.
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                return hresult.S_OK

            def ComMethod(self, arg):
                raise COMError(0, "simulated", (None, None, None, 0, None))

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="WARNING") as logs:
            with self.assertRaises(COMError) as cm:
                obj.ComMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.ComMethod implementation:",
        )
        self.assertEqual(
            logs.records[1].getMessage(),
            "HRESULT_FROM_WIN32(0) is S_OK; using E_FAIL instead so "
            "that the failure is not reported as success",
        )

    def test_comerror_with_unpackable_details_in_call_without_this(self):
        # `call_without_this` unpacks `COMError.details` into five fields to
        # build the error description.  When `details` is not a 5-element
        # sequence, the unpacking raises `ValueError`/`TypeError`; the handler
        # falls back to `str(details)` so the error is still reported.
        class MyObject(COMObject):
            _com_interfaces_ = [self.ITestHresultFromWin32]
            _reg_clsid_ = GUID.create_new()

            def StdMethod(self, this, arg):
                return hresult.S_OK

            def ComMethod(self, arg):
                raise COMError(hresult.E_FAIL, "simulated", (None, None))  # type: ignore

        obj = MyObject().QueryInterface(self.ITestHresultFromWin32)
        with self.assertLogs("comtypes._vtbl", level="ERROR") as logs:
            with self.assertRaises(COMError) as cm:
                obj.ComMethod(0)
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(cm.exception.details, ("(None, None)", None, None, 0, None))
        self.assertEqual(
            logs.records[0].getMessage(),
            "Exception in ITestHresultFromWin32.ComMethod implementation:",
        )
