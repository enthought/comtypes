import ctypes
import unittest as ut
from ctypes import POINTER, byref, pointer

import comtypes
import comtypes.client
from comtypes import IUnknown, hresult
from comtypes.automation import IDispatch

comtypes.client.GetModule("UIAutomationCore.dll")
comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting as scrrun
from comtypes.gen import UIAutomationClient as uiac


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
        ctypes.oledll.ole32.CoCreateInstance(
            byref(scrrun.Dictionary._reg_clsid_),
            None,
            comtypes.CLSCTX_SERVER,
            byref(scrrun.IDictionary._iid_),
            byref(ptr),
        )
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(scrrun.IDictionary._iid_), ptr
        )
        self.assertEqual(hr, hresult.S_OK)

    def test_valid_interface(self):
        dic = POINTER(IDispatch)()
        hr = scrrun.Dictionary().IUnknown_QueryInterface(
            None, pointer(scrrun.IDictionary._iid_), byref(dic)
        )
        self.assertEqual(hr, hresult.S_OK)
        self.assertEqual(dic.GetTypeInfoCount(), 1)  # type: ignore


class Test_IPersist_GetClassID(ut.TestCase):
    def test(self):
        self.assertEqual(
            uiac.CUIAutomation().IPersist_GetClassID(),
            uiac.CUIAutomation._reg_clsid_,
        )
