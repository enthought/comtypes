import unittest
from typing import Any

import comtypes.test.TestComServer
import comtypes.test.TestDispServer
from comtypes.server.register import register, unregister
from comtypes.test.find_memleak import find_memleak

try:
    from win32com.client import Dispatch
    from win32com.client.gencache import EnsureDispatch

    IMPORT_FAILED = False
except ImportError:
    IMPORT_FAILED = True


def setUpModule():
    if IMPORT_FAILED:
        raise unittest.SkipTest("This depends on 'pywin32'.")
    try:
        register(comtypes.test.TestComServer.TestComServer)
        register(comtypes.test.TestDispServer.TestDispServer)
    except WindowsError as e:
        if e.winerror != 5:  # [Error 5] Access is denied
            raise e
        raise unittest.SkipTest(
            "This test requires the tests to be run as admin since it tries to "
            "register the test COM server."
        )


def tearDownModule():
    unregister(comtypes.test.TestDispServer.TestDispServer)
    unregister(comtypes.test.TestComServer.TestComServer)


class BaseServerTest(object):
    def create_object(self) -> Any: ...

    def _find_memleak(self, func):
        bytes = find_memleak(func)
        self.assertFalse(bytes, "Leaks %d bytes" % bytes)  # type: ignore

    def test_get_id(self):
        obj = self.create_object()
        self._find_memleak(lambda: obj.id)

    def test_get_name(self):
        obj = self.create_object()
        self._find_memleak(lambda: obj.name)

    def test_set_name(self):
        obj = self.create_object()

        def func():
            obj.name = "abcde"

        self._find_memleak(func)

    def test_SetName(self):
        obj = self.create_object()

        def func():
            obj.SetName("abcde")

        self._find_memleak(func)

    def test_eval(self):
        obj = self.create_object()

        def func():
            return obj.eval("(1, 2, 3)")

        self.assertEqual(func(), (1, 2, 3))  # type: ignore
        self._find_memleak(func)

    # These tests make no sense with win32com, some of the tests in
    # `test_comserver` are not performed here:
    # `test_get_typeinfo`, `test_getname` and `test_mixedinout`.
    # Not sure about `test_mixedinout`; it raise 'Invalid Number of parameters'
    # Is mixed [in], [out] args not compatible with IDispatch???


class TestInproc(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return Dispatch("TestComServerLib.TestComServer")


class TestLocalServer(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return Dispatch(
            "TestComServerLib.TestComServer", clsctx=comtypes.CLSCTX_LOCAL_SERVER
        )


class Test_win32com_Dispatch(unittest.TestCase):
    @unittest.skip(
        "It likely fails due to bugs in `GenerateChildFromTypeLibSpec` "
        "or `GetModuleForCLSID`."
    )
    def test_win32com_ensure_dispatch(self):
        # EnsureDispatch is case-sensitive
        d = EnsureDispatch("TestDispServerLib.TestDispServer")

        self.assertEqual(d.eval("3.14"), 3.14)
        self.assertEqual(d.eval("1 + 2"), 3)
        self.assertEqual(d.eval("[1 + 2, 'foo', None]"), (3, "foo", None))

        self.assertEqual(d.eval2("3.14"), 3.14)
        self.assertEqual(d.eval2("1 + 2"), 3)
        self.assertEqual(d.eval2("[1 + 2, 'foo', None]"), (3, "foo", None))

        d.eval(
            "__import__('comtypes.client').client.CreateObject('Scripting.Dictionary')"
        )

        server_id = d.eval("id(self)")
        self.assertEqual(d.id, server_id)

        self.assertEqual(d.name, "spam, spam, spam")

        d.SetName("foo bar")
        self.assertEqual(d.name, "foo bar")

        d.name = "blah"
        self.assertEqual(d.name, "blah")

    def test_win32com_dynamic_dispatch(self):
        # dynamic Dispatch is case-IN-sensitive
        d = Dispatch("TestDispServerLib.TestDispServer")

        self.assertEqual(d.eval("3.14"), 3.14)
        self.assertEqual(d.eval("1 + 2"), 3)
        self.assertEqual(d.eval("[1 + 2, 'foo', None]"), (3, "foo", None))

        self.assertEqual(d.eval2("3.14"), 3.14)
        self.assertEqual(d.eval2("1 + 2"), 3)
        self.assertEqual(d.eval2("[1 + 2, 'foo', None]"), (3, "foo", None))

        d.eval(
            "__import__('comtypes.client').client.CreateObject('Scripting.Dictionary')"
        )

        self.assertEqual(d.EVAL("3.14"), 3.14)
        self.assertEqual(d.EVAL("1 + 2"), 3)
        self.assertEqual(d.EVAL("[1 + 2, 'foo', None]"), (3, "foo", None))

        self.assertEqual(d.EVAL2("3.14"), 3.14)
        self.assertEqual(d.EVAL2("1 + 2"), 3)
        self.assertEqual(d.EVAL2("[1 + 2, 'foo', None]"), (3, "foo", None))

        server_id = d.eval("id(self)")
        self.assertEqual(d.id, server_id)
        self.assertEqual(d.ID, server_id)

        self.assertEqual(d.Name, "spam, spam, spam")
        self.assertEqual(d.nAME, "spam, spam, spam")

        d.SetName("foo bar")
        self.assertEqual(d.Name, "foo bar")

        # fails.  Why?
        # d.name = "blah"
        # self.assertEqual(d.Name, "blah")


if __name__ == "__main__":
    unittest.main()
