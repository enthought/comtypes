import sys
import unittest
import comtypes.client
import cStringIO
import ctypes

class EventsTest(unittest.TestCase):
    def setUp(self):
        self.old_stdout = sys.stdout
        sys.stdout = cStringIO.StringIO()

    def tearDown(self):
        sys.stdout = self.old_stdout

    def test(self):
        # Start IE, call .Quit(), and check if the
        # DWebBrowserEvents2_OnQuit event has fired.  We do this by
        # calling ShowEvents() and capturing sys.stdout.
        o = comtypes.client.CreateObject("InternetExplorer.Application")
        conn = comtypes.client.ShowEvents(o)
        o.Quit()
        del o

        comtypes.client.PumpEvents(0.2)

        stream = sys.stdout
        stream.flush()
        sys.stdout = self.old_stdout
        output = stream.getvalue().splitlines()

        self.failUnless('# event found: DWebBrowserEvents2_OnMenuBar' in output)
        self.failUnless('Event DWebBrowserEvents2_OnQuit(None)' in output)

if __name__ == "__main__":
    unittest.main()
