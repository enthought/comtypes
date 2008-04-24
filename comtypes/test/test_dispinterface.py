import unittest

from comtypes.server.register import register#, unregister
from comtypes.test import is_resource_enabled

################################################################
import comtypes.test.TestDispServer
register(comtypes.test.TestDispServer.TestDispServer)

class Test(unittest.TestCase):

    if is_resource_enabled("pythoncom"):
        def test_win32com(self):
            from win32com.client import Dispatch
            d = Dispatch("TestDispServerLib.TestDispServer")

            self.assertEqual(d.eval("3.14"), 3.14)
            self.assertEqual(d.eval("1 + 2"), 3)
            self.assertEqual(d.eval("[1 + 2, 'foo', None]"), (3, 'foo', None))

            self.assertEqual(d.eval2("3.14"), 3.14)
            self.assertEqual(d.eval2("1 + 2"), 3)
            self.assertEqual(d.eval2("[1 + 2, 'foo', None]"), (3, 'foo', None))

            d.eval("__import__('comtypes.client').client.CreateObject('MSScriptControl.ScriptControl')")

            self.assertEqual(d.EVAL("3.14"), 3.14)
            self.assertEqual(d.EVAL("1 + 2"), 3)
            self.assertEqual(d.EVAL("[1 + 2, 'foo', None]"), (3, 'foo', None))

            self.assertEqual(d.EVAL2("3.14"), 3.14)
            self.assertEqual(d.EVAL2("1 + 2"), 3)
            self.assertEqual(d.EVAL2("[1 + 2, 'foo', None]"), (3, 'foo', None))

            server_id = d.eval("id(self)")
            self.assertEqual(d.id, server_id)
            self.assertEqual(d.ID, server_id)

            self.assertEqual(d.Name, "spam, spam, spam")
            self.assertEqual(d.nAME, "spam, spam, spam")

            d.SetName("foo bar")
            self.assertEqual(d.Name, "foo bar")

            # fails
##            d.name = "blah"
##            self.assertEqual(d.Name, "blah")

        def test_win32com_dyndispatch(self):
            from win32com.client.dynamic import Dispatch
            d = Dispatch("TestDispServerLib.TestDispServer")

            self.assertEqual(d.eval("3.14"), 3.14)
            self.assertEqual(d.eval("1 + 2"), 3)
            self.assertEqual(d.eval("[1 + 2, 'foo', None]"), (3, 'foo', None))

            self.assertEqual(d.eval2("3.14"), 3.14)
            self.assertEqual(d.eval2("1 + 2"), 3)
            self.assertEqual(d.eval2("[1 + 2, 'foo', None]"), (3, 'foo', None))

            d.eval("__import__('comtypes.client').client.CreateObject('MSScriptControl.ScriptControl')")

            self.assertEqual(d.EVAL("3.14"), 3.14)
            self.assertEqual(d.EVAL("1 + 2"), 3)
            self.assertEqual(d.EVAL("[1 + 2, 'foo', None]"), (3, 'foo', None))

            self.assertEqual(d.EVAL2("3.14"), 3.14)
            self.assertEqual(d.EVAL2("1 + 2"), 3)
            self.assertEqual(d.EVAL2("[1 + 2, 'foo', None]"), (3, 'foo', None))

            server_id = d.eval("id(self)")
            self.assertEqual(d.id, server_id)
            self.assertEqual(d.ID, server_id)

            self.assertEqual(d.Name, "spam, spam, spam")
            self.assertEqual(d.nAME, "spam, spam, spam")

            d.SetName("foo bar")
            self.assertEqual(d.Name, "foo bar")

            # fails
##            d.name = "blah"
##            self.assertEqual(d.Name, "blah")

    def test_comtypes(self):
        from comtypes.client import CreateObject
        d = CreateObject("TestDispServerLib.TestDispServer")

        self.assertEqual(d.eval("3.14"), 3.14)
        self.assertEqual(d.eval("1 + 2"), 3)
        self.assertEqual(d.eval("[1 + 2, 'foo', None]"), (3, 'foo', None))

        self.assertEqual(d.eval2("3.14"), 3.14)
        self.assertEqual(d.eval2("1 + 2"), 3)
        self.assertEqual(d.eval2("[1 + 2, 'foo', None]"), (3, 'foo', None))

        d.eval("__import__('comtypes.client').client.CreateObject('MSScriptControl.ScriptControl')")

        self.assertEqual(d.EVAL("3.14"), 3.14)
        self.assertEqual(d.EVAL("1 + 2"), 3)
        self.assertEqual(d.EVAL("[1 + 2, 'foo', None]"), (3, 'foo', None))

        self.assertEqual(d.EVAL2("3.14"), 3.14)
        self.assertEqual(d.EVAL2("1 + 2"), 3)
        self.assertEqual(d.EVAL2("[1 + 2, 'foo', None]"), (3, 'foo', None))

        server_id = d.eval("id(self)")
        self.assertEqual(d.id, server_id)
        self.assertEqual(d.ID, server_id)

        self.assertEqual(d.Name, "spam, spam, spam")
        self.assertEqual(d.nAME, "spam, spam, spam")

        d.SetName("foo bar")
        self.assertEqual(d.Name, "foo bar")

        # fails
##        d.name = "blah"
##        self.assertEqual(d.Name, "blah")

if __name__ == "__main__":
    unittest.main()
