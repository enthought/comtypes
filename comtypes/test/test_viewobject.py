import unittest

import comtypes.client
from comtypes import IUnknown
from comtypes.viewobject import DVASPECT_CONTENT, IAdviseSink, IViewObject


def create_shell_explorer() -> IUnknown:
    return comtypes.client.CreateObject("Shell.Explorer")


class Test_IViewObject(unittest.TestCase):
    def test_Advise_GetAdvise(self):
        vo = create_shell_explorer().QueryInterface(IViewObject)
        # Test that we can clear any existing advise connection.
        vo.SetAdvise(DVASPECT_CONTENT, 0, None)
        # Verify that no advise connection is present.
        aspect, advf, sink = vo.GetAdvise()
        self.assertIsInstance(aspect, int)
        self.assertIsInstance(advf, int)
        self.assertIsInstance(sink, IAdviseSink)
        self.assertFalse(sink)  # A NULL com pointer evaluates to False.

    def test_Freeze_Unfreeze(self):
        vo = create_shell_explorer().QueryInterface(IViewObject)
        cookie = vo.Freeze(DVASPECT_CONTENT, -1, None)
        self.assertIsInstance(cookie, int)
        vo.Unfreeze(cookie)
