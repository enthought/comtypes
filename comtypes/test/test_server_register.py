import _ctypes
import ctypes
import os
import sys
import unittest as ut
import winreg
from unittest import mock

import comtypes
from comtypes import GUID
from comtypes.server import register
from comtypes.server.register import RegistryEntries, _get_serverdll

HKCR = winreg.HKEY_CLASSES_ROOT


class Test_get_serverdll(ut.TestCase):
    def test_nonfrozen(self):
        self.assertEqual(_ctypes.__file__, _get_serverdll())

    def test_frozen(self):
        with mock.patch.object(register, "sys") as _sys:
            with mock.patch.object(register, "windll") as _windll:
                handle = 1234
                _sys.frozendllhandle = handle
                self.assertEqual(b"\x00" * 260, _get_serverdll())
                GetModuleFileName = _windll.kernel32.GetModuleFileNameA
                (((hModule, lpFilename, nSize), _),) = GetModuleFileName.call_args_list
                self.assertEqual(handle, hModule)
                buf_type = type(ctypes.create_string_buffer(260))
                self.assertIsInstance(lpFilename, buf_type)
                self.assertEqual(260, nSize)


class Test_RegistryEntries_NonFrozen(ut.TestCase):
    def test_reg_clsid(self):
        reg_clsid = GUID.create_new()

        class Cls:
            _reg_clsid_ = reg_clsid

        expected = [(HKCR, rf"CLSID\{reg_clsid}", "", "")]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_reg_desc(self):
        reg_clsid = GUID.create_new()
        reg_desc = "description for testing"

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_desc_ = reg_desc

        expected = [(HKCR, rf"CLSID\{reg_clsid}", "", reg_desc)]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_reg_novers_progid(self):
        reg_clsid = GUID.create_new()
        reg_novers_progid = "Lib.Server"

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_novers_progid_ = reg_novers_progid

        expected = [(HKCR, rf"CLSID\{reg_clsid}", "", "Lib Server")]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_progid(self):
        reg_clsid = GUID.create_new()
        reg_progid = "Lib.Server.1"

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_progid_ = reg_progid

        expected = [
            (HKCR, rf"CLSID\{reg_clsid}", "", "Lib Server 1"),
            (HKCR, rf"CLSID\{reg_clsid}\ProgID", "", reg_progid),
            (HKCR, reg_progid, "", "Lib Server 1"),
            (HKCR, rf"{reg_progid}\CLSID", "", str(reg_clsid)),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_reg_progid_reg_desc(self):
        reg_clsid = GUID.create_new()
        reg_progid = "Lib.Server.1"
        reg_desc = "description for testing"

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_progid_ = reg_progid
            _reg_desc_ = reg_desc

        expected = [
            (HKCR, rf"CLSID\{reg_clsid}", "", reg_desc),
            (HKCR, rf"CLSID\{reg_clsid}\ProgID", "", reg_progid),
            (HKCR, reg_progid, "", "description for testing"),
            (HKCR, rf"{reg_progid}\CLSID", "", str(reg_clsid)),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_reg_progid_reg_novers_progid(self):
        reg_clsid = GUID.create_new()
        reg_progid = "Lib.Server.1"
        reg_novers_progid = "Lib.Server"

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_progid_ = reg_progid
            _reg_novers_progid_ = reg_novers_progid

        clsid_sub = rf"CLSID\{reg_clsid}"
        expected = [
            (HKCR, clsid_sub, "", "Lib Server"),
            (HKCR, rf"{clsid_sub}\ProgID", "", reg_progid),
            (HKCR, reg_progid, "", "Lib Server"),
            (HKCR, rf"{reg_progid}\CLSID", "", str(reg_clsid)),
            (HKCR, rf"{clsid_sub}\VersionIndependentProgID", "", reg_novers_progid),
            (HKCR, reg_novers_progid, "", "Lib Server"),
            (HKCR, rf"{reg_novers_progid}\CurVer", "", "Lib.Server.1"),
            (HKCR, rf"{reg_novers_progid}\CLSID", "", str(reg_clsid)),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_local_server(self):
        reg_clsid = GUID.create_new()
        reg_clsctx = comtypes.CLSCTX_LOCAL_SERVER

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_clsctx_ = reg_clsctx

        clsid_sub = rf"CLSID\{reg_clsid}"
        local_srv_sub = rf"{clsid_sub}\LocalServer32"
        expected = [
            (HKCR, clsid_sub, "", ""),
            (HKCR, local_srv_sub, "", f"{sys.executable} {__file__}"),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_inproc_server(self):
        reg_clsid = GUID.create_new()
        reg_clsctx = comtypes.CLSCTX_INPROC_SERVER

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_clsctx_ = reg_clsctx

        clsid_sub = rf"CLSID\{reg_clsid}"
        inproc_srv_sub = rf"{clsid_sub}\InprocServer32"
        full_classname = f"{__name__}.Cls"
        expected = [
            (HKCR, clsid_sub, "", ""),
            (HKCR, inproc_srv_sub, "", _ctypes.__file__),
            (HKCR, inproc_srv_sub, "PythonClass", full_classname),
            (HKCR, inproc_srv_sub, "PythonPath", os.path.dirname(__file__)),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_inproc_server_reg_threading(self):
        reg_clsid = GUID.create_new()
        reg_threading = "Both"
        reg_clsctx = comtypes.CLSCTX_INPROC_SERVER

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_threading_ = reg_threading
            _reg_clsctx_ = reg_clsctx

        clsid_sub = rf"CLSID\{reg_clsid}"
        inproc_srv_sub = rf"{clsid_sub}\InprocServer32"
        full_classname = f"{__name__}.Cls"
        expected = [
            (HKCR, clsid_sub, "", ""),
            (HKCR, inproc_srv_sub, "", _ctypes.__file__),
            (HKCR, inproc_srv_sub, "PythonClass", full_classname),
            (HKCR, inproc_srv_sub, "PythonPath", os.path.dirname(__file__)),
            (HKCR, inproc_srv_sub, "ThreadingModel", reg_threading),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_reg_typelib(self):
        reg_clsid = GUID.create_new()
        libid = str(GUID.create_new())
        reg_typelib = (libid, 1, 0)

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_typelib_ = reg_typelib

        expected = [
            (HKCR, rf"CLSID\{reg_clsid}", "", ""),
            (HKCR, rf"CLSID\{reg_clsid}\Typelib", "", libid),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))

    def test_all_entries(self):
        reg_clsid = GUID.create_new()
        libid = str(GUID.create_new())
        reg_typelib = (libid, 1, 0)
        reg_threading = "Both"
        reg_progid = "Lib.Server.1"
        reg_novers_progid = "Lib.Server"
        reg_desc = "description for testing"
        reg_clsctx = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER

        class Cls:
            _reg_clsid_ = reg_clsid
            _reg_typelib_ = reg_typelib
            _reg_threading_ = reg_threading
            _reg_progid_ = reg_progid
            _reg_novers_progid_ = reg_novers_progid
            _reg_desc_ = reg_desc
            _reg_clsctx_ = reg_clsctx

        clsid_sub = rf"CLSID\{reg_clsid}"
        inproc_srv_sub = rf"{clsid_sub}\InprocServer32"
        local_srv_sub = rf"{clsid_sub}\LocalServer32"
        full_classname = f"{__name__}.Cls"
        expected = [
            (HKCR, clsid_sub, "", reg_desc),
            (HKCR, rf"{clsid_sub}\ProgID", "", reg_progid),
            (HKCR, reg_progid, "", reg_desc),
            (HKCR, rf"{reg_progid}\CLSID", "", str(reg_clsid)),
            (HKCR, rf"{clsid_sub}\VersionIndependentProgID", "", reg_novers_progid),
            (HKCR, reg_novers_progid, "", reg_desc),
            (HKCR, rf"{reg_novers_progid}\CurVer", "", reg_progid),
            (HKCR, rf"{reg_novers_progid}\CLSID", "", str(reg_clsid)),
            (HKCR, local_srv_sub, "", f"{sys.executable} {__file__}"),
            (HKCR, inproc_srv_sub, "", _ctypes.__file__),
            (HKCR, inproc_srv_sub, "PythonClass", full_classname),
            (HKCR, inproc_srv_sub, "PythonPath", os.path.dirname(__file__)),
            (HKCR, inproc_srv_sub, "ThreadingModel", reg_threading),
            (HKCR, rf"{clsid_sub}\Typelib", "", libid),
        ]
        self.assertEqual(expected, list(RegistryEntries(Cls)))
