import ctypes
import unittest as ut

from comtypes import IUnknown, automation
from comtypes.client import CreateObject, GetModule, dynamic, lazybind


class Test_Dispatch_Function(ut.TestCase):
    def test_returns_lazybind_Dispatch(self):
        # When `dynamic=True`, objects providing type information will return a
        # `lazybind.Dispatch` instance.
        orig = CreateObject("Scripting.Dictionary", interface=automation.IDispatch)
        disp = dynamic.Dispatch(orig)
        self.assertIsInstance(disp, lazybind.Dispatch)
        # Calling `dynamic.Dispatch` with an already dispatched object should
        # return the same instance.
        self.assertIs(disp, dynamic.Dispatch(disp))

    def test_returns_dynamic_Dispatch(self):
        # When `dynamic=True`, objects that do NOT provide type information (or
        # fail to provide it) will return a `dynamic._Dispatch` instance.
        orig = CreateObject(
            "WindowsInstaller.Installer", interface=automation.IDispatch
        )
        disp = dynamic.Dispatch(orig)
        self.assertIsInstance(disp, dynamic._Dispatch)
        # Calling `dynamic.Dispatch` on an already dispatched object should
        # return the same instance.
        self.assertIs(disp, dynamic.Dispatch(disp))


class Test_Dispatch_Class(ut.TestCase):
    # `MethodCaller` and `_Collection` are indirectly covered in this.
    def test_dict(self):
        # The following conditional branches are not covered;
        # - not `hresult in ERRORS_BAD_CONTEXT`
        # - not `0 != enum.Skip(index)`
        # - other than `COMError` raises in `__getattr__`
        orig = CreateObject("Scripting.Dictionary", interface=automation.IDispatch)
        d = dynamic._Dispatch(orig)
        d.CompareMode = 42
        d.Item["foo"] = 1
        d.Item["bar"] = "spam foo"
        d.Item["baz"] = 3.14
        self.assertEqual(d[0], "foo")
        self.assertEqual(d.Item["foo"], 1)
        self.assertEqual([k for k in iter(d)], ["foo", "bar", "baz"])
        self.assertIsInstance(hash(d), int)
        d._FlagAsMethod("_NewEnum")
        self.assertIs(type(d._NewEnum()), ctypes.POINTER(IUnknown))
        scrrun = GetModule("scrrun.dll")
        scr_dict = d.QueryInterface(scrrun.IDictionary)
        self.assertIsInstance(scr_dict, scrrun.IDictionary)
        d.Item["qux"] = scr_dict
        with self.assertRaises(IndexError):
            d[4]
        with self.assertRaises(AttributeError):
            d.__foo__


if __name__ == "__main__":
    ut.main()
