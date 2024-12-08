import os
import unittest

import comtypes.test.TestDispServer
from comtypes.server.register import register, unregister


def setUpModule():
    try:
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


class Test_comtypes(unittest.TestCase):
    def test_comtypes(self):
        from comtypes.client import CreateObject

        d = CreateObject("TestDispServerLib.TestDispServer")

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

        d.name = "blah"
        self.assertEqual(d.Name, "blah")


@unittest.skip("This raises 'ClassFactory cannot supply requested class'. Why?")
class Test_jscript(unittest.TestCase):
    def test_withjscript(self):
        jscript = os.path.join(os.path.dirname(__file__), "test_jscript.js")
        errcode = os.system("cscript -nologo %s" % jscript)
        self.assertEqual(errcode, 0)


if __name__ == "__main__":
    unittest.main()
