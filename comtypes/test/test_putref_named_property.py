import unittest

from comtypes import CLSCTX_LOCAL_SERVER, IUnknown
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
class TestComtypesDispPutRef(unittest.TestCase):
    """`Item` reached through `IDispatch::Invoke` (the dispinterface alias)."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.ComtypesDispPutRefTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispPutRefTest,
        )

    def test_propputref_setter_stores_zero_sized_obj(self):
        other = CreateObject("Scripting.Dictionary")
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )

    def test_propputref_setter_stores_non_zero_sized_obj(self):
        other = CreateObject("Scripting.Dictionary")
        other["foo"] = "bar"
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )

    def test_propputref_setter_stores_non_sized_obj(self):
        other = CreateObject("StdFont")
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestComtypesDualPutRef(unittest.TestCase):
    """`Item` reached through the dual vtable interface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.ComtypesDispPutRefTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualPutRefTest,
        )

    def test_propputref_setter_stores_zero_sized_obj(self):
        other = CreateObject("Scripting.Dictionary")
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )

    def test_propputref_setter_stores_non_zero_sized_obj(self):
        other = CreateObject("Scripting.Dictionary")
        other["foo"] = "bar"
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )

    def test_propputref_setter_stores_non_sized_obj(self):
        other = CreateObject("StdFont")
        self.obj.Item[3] = other
        result = self.obj.Item[3]
        self.assertEqual(
            result.QueryInterface(IUnknown), other.QueryInterface(IUnknown)
        )


if __name__ == "__main__":
    unittest.main()
