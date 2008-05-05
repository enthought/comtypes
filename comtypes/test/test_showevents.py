import sys
import unittest
import comtypes.test.test_showevents
import doctest

class EventsTest(unittest.TestCase):

    def test(self):
        doctest.testmod(comtypes.test.test_showevents, optionflags=doctest.ELLIPSIS)

    # ShowEvents and GetEvents are never called, they only contain doctests:
    def ShowEvents(self):
        '''
        >>> from comtypes.client import CreateObject, ShowEvents, PumpEvents
        >>>
        >>> o = CreateObject("InternetExplorer.Application")
        >>> con = ShowEvents(o)
        # event found: DWebBrowserEvents2_StatusTextChange
        # event found: DWebBrowserEvents2_ProgressChange
        # event found: DWebBrowserEvents2_CommandStateChange
        # event found: DWebBrowserEvents2_DownloadBegin
        # event found: DWebBrowserEvents2_DownloadComplete
        # event found: DWebBrowserEvents2_TitleChange
        # event found: DWebBrowserEvents2_PropertyChange
        # event found: DWebBrowserEvents2_BeforeNavigate2
        # event found: DWebBrowserEvents2_NewWindow2
        # event found: DWebBrowserEvents2_NavigateComplete2
        # event found: DWebBrowserEvents2_DocumentComplete
        # event found: DWebBrowserEvents2_OnQuit
        # event found: DWebBrowserEvents2_OnVisible
        # event found: DWebBrowserEvents2_OnToolBar
        # event found: DWebBrowserEvents2_OnMenuBar
        # event found: DWebBrowserEvents2_OnStatusBar
        # event found: DWebBrowserEvents2_OnFullScreen
        # event found: DWebBrowserEvents2_OnTheaterMode
        # event found: DWebBrowserEvents2_WindowSetResizable
        # event found: DWebBrowserEvents2_WindowSetLeft
        # event found: DWebBrowserEvents2_WindowSetTop
        # event found: DWebBrowserEvents2_WindowSetWidth
        # event found: DWebBrowserEvents2_WindowSetHeight
        # event found: DWebBrowserEvents2_WindowClosing
        # event found: DWebBrowserEvents2_ClientToHostWindow
        # event found: DWebBrowserEvents2_SetSecureLockIcon
        # event found: DWebBrowserEvents2_FileDownload
        # event found: DWebBrowserEvents2_NavigateError
        # event found: DWebBrowserEvents2_PrintTemplateInstantiation
        # event found: DWebBrowserEvents2_PrintTemplateTeardown
        # event found: DWebBrowserEvents2_UpdatePageStatus
        # event found: DWebBrowserEvents2_PrivacyImpactedStateChange
        # event found: DWebBrowserEvents2_NewWindow3
        # event found: DWebBrowserEvents2_SetPhishingFilterStatus
        # event found: DWebBrowserEvents2_WindowStateChanged
        >>> res = o.Navigate2("http://www.python.org")
        Event DWebBrowserEvents2_PropertyChange(None, u'{265b75c1-4158-11d0-90f6-00c04fd497ea}')
        Event DWebBrowserEvents2_BeforeNavigate2(None, <POINTER(IWebBrowser2) ptr=... at ...>, u'http://www.python.org/', 0, None, None, None, False)
        Event DWebBrowserEvents2_DownloadBegin(None)
        Event DWebBrowserEvents2_PropertyChange(None, u'{D0FCA420-D3F5-11CF-B211-00AA004AE837}')
        >>> res = PumpEvents(0.01)
        Event DWebBrowserEvents2_CommandStateChange(None, 2, False)
        Event DWebBrowserEvents2_CommandStateChange(None, 1, False)
        >>> res = o.Quit()
        >>> res = PumpEvents(0.01)
        Event DWebBrowserEvents2_OnQuit(None)
        >>>
        '''

    def GetEvents():
        """
        >>> from comtypes.client import CreateObject, GetEvents, PumpEvents
        >>>
        >>> o =  CreateObject("InternetExplorer.Application")
        >>> class EventHandler(object):
        ...     def DWebBrowserEvents2_PropertyChange(self, this, what):
        ...         print "PropertyChange:", what
        ...         return 0
        ...
        >>>
        >>> con = GetEvents(o, EventHandler())
        >>> res = o.Navigate2("http://www.python.org")
        PropertyChange: {265b75c1-4158-11d0-90f6-00c04fd497ea}
        PropertyChange: {D0FCA420-D3F5-11CF-B211-00AA004AE837}
        >>> res = o.Quit()
        >>> res = PumpEvents(0.01)
        >>>
        """

if __name__ == "__main__":
    unittest.main()
