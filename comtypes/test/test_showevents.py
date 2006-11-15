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
        o = comtypes.client.CreateObject("InternetExplorer.Application")
        conn = comtypes.client.ShowEvents(o)
        o.Quit()
        del o

        # The following code waits for 'timeout' milliseconds in the
        # way required for COM, internally doing the correct things
        # depending on the COM appartment of the current thread.
        timeout = 1000
        # We have to supply at least one NULL handle, otherwise
        # CoWaitForMultipleHandles complains.
        handles = (ctypes.c_void_p * 1)()
        RPC_S_CALLPENDING = -2147417835
        try:
            ctypes.oledll.ole32.CoWaitForMultipleHandles(0,
                                                         timeout,
                                                         len(handles), handles,
                                                         ctypes.byref(ctypes.c_ulong()))
        except WindowsError, details:
            if details[0] != RPC_S_CALLPENDING: # timeout expired
                raise

        stream = sys.stdout
        stream.flush()
        sys.stdout = self.old_stdout
        output = stream.getvalue().splitlines()

        self.failUnless('# event found: DWebBrowserEvents2_OnMenuBar' in output)
        self.failUnless('Event DWebBrowserEvents2_OnQuit(None)' in output)

if __name__ == "__main__":
    unittest.main()
