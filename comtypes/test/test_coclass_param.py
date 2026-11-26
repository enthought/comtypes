"""Tests for a method whose declared return type is a coclass."""

import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDualCoClassParam(unittest.TestCase):
    """A coclass return type reached through the dual vtable interface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.CoClassParamTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualCoClassParamTest,
        )

    def test_result_returns_default_interface_of_coclass(self):
        result = self.obj.GetItemAsCC()
        self.assertIsInstance(result, ComtypesCppTestSrvLib.ICoClassParamTestItem)
        self.assertEqual(result.GetValue(), 42)

    def test_result_interface_returns_interface(self):
        result = self.obj.GetItemAsItf()
        self.assertIsInstance(result, ComtypesCppTestSrvLib.ICoClassParamTestItem)
        self.assertEqual(result.GetValue(), 42)

    def test_self_as_itf_and_self_as_cc_return_the_same_object(self):
        item = self.obj.GetItemAsItf()
        self.assertEqual(item.SelfAsItf, item.SelfAsCC)
        self.assertIsInstance(
            item.SelfAsItf, ComtypesCppTestSrvLib.ICoClassParamTestItem
        )
        self.assertIsInstance(
            item.SelfAsCC, ComtypesCppTestSrvLib.ICoClassParamTestItem
        )
        self.assertEqual(item, item.SelfAsItf)
        self.assertEqual(item, item.SelfAsCC)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDispCoClassParam(unittest.TestCase):
    """A coclass return type reached through IDispatch::Invoke."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.CoClassParamTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispCoClassParamTest,
        )

    def test_result_returns_raw_coclass_pointer(self):
        result = self.obj.GetItemAsCC()
        self.assertNotIsInstance(
            result,
            ComtypesCppTestSrvLib.ICoClassParamTestItem,
            msg="CHARACTERIZATION: failure is expected; the dispinterface "
            "declares the return type as the `CoCoClassParamTestItem` coclass, "
            "so the result should be `QueryInterface`'d to the coclass' default "
            "interface, but comtypes generates `POINTER(CoCoClassParamTestItem)` "
            "which has no `__com_interface__`, so the raw `IUnknown` pointer "
            "is returned instead; once the dispinterface method `QueryInterface`s "
            "a coclass return value, `result` should be an `ICoClassParamTestItem`.",
        )

    def test_result_interface_returns_interface(self):
        result = self.obj.GetItemAsItf()
        self.assertIsInstance(result, ComtypesCppTestSrvLib.ICoClassParamTestItem)
        self.assertEqual(result.GetValue(), 42)

    def test_self_as_itf_and_self_as_cc_return_the_same_object(self):
        item = self.obj.GetItemAsItf()
        self.assertEqual(item.SelfAsItf, item.SelfAsCC)
        self.assertIsInstance(
            item.SelfAsItf, ComtypesCppTestSrvLib.ICoClassParamTestItem
        )
        self.assertIsInstance(
            item.SelfAsCC, ComtypesCppTestSrvLib.ICoClassParamTestItem
        )
        self.assertEqual(item, item.SelfAsItf)
        self.assertEqual(item, item.SelfAsCC)
