import sys
import unittest as ut
from ctypes import POINTER

import comtypes.client
from comtypes import CLSCTX_INPROC_SERVER

if sys.version_info >= (3, 11):
    from typing import assert_type
else:
    from typing_extensions import assert_type

# Generate/ensure the existence of the `UIAutomationClient` module.
comtypes.client.GetModule("UIAutomationCore.dll")
from comtypes.gen.UIAutomationClient import (
    CUIAutomation,
    IUIAutomation,
    IUIAutomationElement,
)


class TestIUIAutomation(ut.TestCase):
    def test(self):
        # Create an instance of the `CUIAutomation` coclass requesting the
        # `IUIAutomation` interface.
        # `CreateObject` returns an instance typed as the requested `interface`
        # (`IUIAutomation`), while at runtime it is a `POINTER(IUIAutomation)`
        # that behaves as an `IUIAutomation` subclass via its metaclass.
        iuia = comtypes.client.CreateObject(
            CUIAutomation().IPersist_GetClassID(),
            interface=IUIAutomation,
            clsctx=CLSCTX_INPROC_SERVER,
        )
        self.assertIsInstance(iuia, POINTER(IUIAutomation))
        self.assertIsInstance(iuia, IUIAutomation)
        assert_type(iuia, IUIAutomation)
        # `GetRootElement` returns a COM pointer to the root `IUIAutomationElement`.
        # Verify that both runtime type and static type inference resolve to
        # `IUIAutomationElement`.
        root = iuia.GetRootElement()
        self.assertIsInstance(root, POINTER(IUIAutomationElement))
        self.assertIsInstance(root, IUIAutomationElement)
        assert_type(root, IUIAutomationElement)
        # Unlike `IDictionary` or collection interfaces, standard COM interfaces
        # like `IUIAutomation` are neither subscriptable nor callable.
        # Verify that indexing and calling raise TypeError at runtime, and are
        # flagged by static type checkers (requiring type: ignore).
        with self.assertRaises(TypeError):
            iuia[1]  # type: ignore
        with self.assertRaises(TypeError):
            iuia(1)  # type: ignore
