import unittest
from ctypes import *
from ctypes.wintypes import *
from comtypes.client import CreateObject
from comtypes.server.register import register, unregister
from comtypes.test import is_resource_enabled
from comtypes.test.find_memleak import find_memleak

try:
    any
except NameError:
    from comtypes.test.find_memleak import any

################################################################
import comtypes.test.TestServer

LOOPS = 10, 1000

class TestInproc(unittest.TestCase):

    def __init__(self, *args, **kw):
        register(comtypes.test.TestServer.TestComServer)
        super(TestInproc, self).__init__(*args, **kw)

    def create_object(self):
        return CreateObject("TestComServerLib.TestComServer")

    def _find_memleak(self, func):
        leaks = find_memleak(func)
        self.failIf(any(leaks), "Leaks %d bytes: %s" % (sum(leaks), leaks))

    if is_resource_enabled("memleaks"):
        def test_get_id(self):
            obj = self.create_object()
            self._find_memleak(lambda: obj.id)

        def test_get_name(self):
            obj = self.create_object()
            self._find_memleak(lambda: obj.name)

        # This leaks memory, but only with comtypes client code,
        # not win32com client code
        def test_set_name(self):
            obj = self.create_object()
            def func():
                obj.name = u"abcde"
            self._find_memleak(func)

        def test_SetName(self):
            obj = self.create_object()
            def func():
                obj.SetName(u"abcde")
            self._find_memleak(func)


        def test_eval(self):
            obj = self.create_object()
            def func():
                obj.eval("(1, 2, 3)")
            self._find_memleak(func)

        def test_get_typeinfo(self):
            obj = self.create_object()
            def func():
                obj.GetTypeInfo(0)
                obj.GetTypeInfoCount()
                obj.QueryInterface(comtypes.IUnknown)
            self._find_memleak(func)

class TestLocalServer(TestInproc):
    def create_object(self):
        return CreateObject("TestComServerLib.TestComServer",
                            clsctx = comtypes.CLSCTX_LOCAL_SERVER)

try:
    from win32com.client import Dispatch
except ImportError:
    pass
else:
    class TestInproc_win32com(TestInproc):
        def create_object(self):
            return Dispatch("TestComServerLib.TestComServer")

        def test_get_typeinfo(self):
            pass

    class TestLocalServer_win32com(TestInproc):
        def create_object(self):
            return Dispatch("TestComServerLib.TestComServer", clsctx = comtypes.CLSCTX_LOCAL_SERVER)

        def test_get_typeinfo(self):
            pass

if __name__ == "__main__":
    unittest.main()
