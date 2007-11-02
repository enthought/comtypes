import sys
import unittest
import comtypes.client
import cStringIO
import ctypes

# XXX Move this into comtypes.client._event, and make it available
# from comtyes.client. After finding a decent name for it.
def pump_messages(timeout):
    # This following code waits for 'timeout' milliseconds in the way
    # required for COM, internally doing the correct things depending
    # on the COM appartment of the current thread.  It is possible to
    # terminate the message loop by pressing CTRL+C, which will raise
    # a KeyboardInterrupt.
    hevt = ctypes.windll.kernel32.CreateEventA(None, True, False, None)
    handles = (ctypes.c_void_p * 1)(hevt)
    RPC_S_CALLPENDING = -2147417835

    @ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_uint)
    def HandlerRoutine(dwCtrlType):
        if dwCtrlType == 0: # CTRL+C
            ctypes.windll.kernel32.SetEvent(hevt)
            return 1
        return 0

    ctypes.windll.kernel32.SetConsoleCtrlHandler(HandlerRoutine, 1)

    try:
        try:
            res = ctypes.oledll.ole32.CoWaitForMultipleHandles(0,
                                                               int(timeout * 1000),
                                                               len(handles), handles,
                                                               ctypes.byref(ctypes.c_ulong()))
        except WindowsError, details:
            if details[0] != RPC_S_CALLPENDING: # timeout expired
                raise
        else:
            raise KeyboardInterrupt
    finally:
        ctypes.windll.kernel32.CloseHandle(hevt)
        ctypes.windll.kernel32.SetConsoleCtrlHandler(HandlerRoutine, 0)

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

        pump_messages(0.2)

        stream = sys.stdout
        stream.flush()
        sys.stdout = self.old_stdout
        output = stream.getvalue().splitlines()

        self.failUnless('# event found: DWebBrowserEvents2_OnMenuBar' in output)
        self.failUnless('Event DWebBrowserEvents2_OnQuit(None)' in output)

if __name__ == "__main__":
    unittest.main()
