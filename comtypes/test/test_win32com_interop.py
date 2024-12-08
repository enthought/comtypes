import unittest
from ctypes import POINTER, PyDLL, byref, c_void_p, py_object
from ctypes.wintypes import BOOL

from comtypes import IUnknown
from comtypes.automation import IDispatch
from comtypes.client import CreateObject

try:
    import pythoncom
    import win32com.client

    IMPORT_FAILED = False
    # We use the PyCom_PyObjectFromIUnknown function in pythoncom25.dll to
    # convert a comtypes COM pointer into a pythoncom COM pointer.
    # Fortunately this function is exported by the dll...
    #
    # This is the C prototype; we must pass 'True' as third argument:
    #
    # PyObject *PyCom_PyObjectFromIUnknown(IUnknown *punk, REFIID riid, BOOL bAddRef)

    _PyCom_PyObjectFromIUnknown = PyDLL(pythoncom.__file__).PyCom_PyObjectFromIUnknown
    _PyCom_PyObjectFromIUnknown.restype = py_object
    _PyCom_PyObjectFromIUnknown.argtypes = (POINTER(IUnknown), c_void_p, BOOL)
except ImportError:
    # this test depends on pythoncom but it is not available.  Maybe we should just skip it even
    # if it is available since pythoncom is not a project dependency and adding tests depending
    # on the vagaries of various testing environments is not deterministic.
    # TODO: Evaluate if we should just remove this test or what.
    IMPORT_FAILED = True


def setUpModule():
    if IMPORT_FAILED:
        raise unittest.SkipTest(
            "This test requires the pythoncom library installed.  If this is "
            "important tests then we need to add dev dependencies to the project that include pythoncom."
        )


def comtypes2pywin(ptr, interface=None):
    """Convert a comtypes pointer 'ptr' into a pythoncom
    PyI<interface> object.

    'interface' specifies the interface we want; it must be a comtypes
    interface class.  The interface must be implemented by the object;
    and the interface must be known to pythoncom.

    If 'interface' is specified, comtypes.IUnknown is used.
    """
    if interface is None:
        interface = IUnknown
    return _PyCom_PyObjectFromIUnknown(ptr, byref(interface._iid_), True)


################################################################


def comtypes_get_refcount(ptr):
    """Helper function for testing: return the COM reference count of
    a comtypes COM object"""
    ptr.AddRef()
    return ptr.Release()


from comtypes import COMObject


class MyComObject(COMObject):
    """A completely trivial COM object implementing IDispatch. Calling
    any methods will return the error code E_NOTIMPL (except the
    IUnknown methods; they are implemented in the base class."""

    _com_interfaces_ = [IDispatch]


################################################################


class ConvertComtypesPtrToPythonComObjTest(unittest.TestCase):
    def test_mycomobject(self):
        o = MyComObject()
        p = comtypes2pywin(o, IDispatch)
        disp = win32com.client.Dispatch(p)
        self.assertEqual(repr(disp), "<COMObject <unknown>>")

    def test_refcount(self):
        # Convert a comtypes COM interface pointer into a win32com COM pointer.
        dic = CreateObject("Scripting.Dictionary")
        # The COM refcount of the created object is 1:
        self.assertEqual(comtypes_get_refcount(dic), 1)
        dic["foo"] = "bar"
        self.assertEqual(dic.Item("foo"), "bar")

        # Create a pythoncom PyIDispatch object from it:
        p = comtypes2pywin(dic, interface=IDispatch)
        self.assertEqual(comtypes_get_refcount(dic), 2)

        # Make it usable...
        disp = win32com.client.Dispatch(p)
        self.assertEqual(comtypes_get_refcount(dic), 2)
        self.assertEqual(disp.Item("foo"), "bar")

        # Cleanup and make sure that the COM refcounts are correct
        del p, disp
        self.assertEqual(comtypes_get_refcount(dic), 1)


if __name__ == "__main__":
    unittest.main()
