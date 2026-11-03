import sys
import unittest as ut
from ctypes import POINTER

import comtypes.client

if sys.version_info >= (3, 11):
    from typing import assert_type
else:
    from typing_extensions import assert_type

# Generate/ensure the existence of the `Scripting` module.
comtypes.client.GetModule("scrrun.dll")
# Must be imported statically; otherwise it will not be analyzed statically.
from comtypes.gen import Scripting


class TestScriptingDictionary(ut.TestCase):
    def test(self):
        # `CreateObject` is a generic function that returns an instance of the
        # `IUnknown` subclass specified by the `interface` argument.
        # Although the actual runtime type is `POINTER(IDictionary)`, it behaves
        # as a subclass of `IDictionary` via its metaclass.
        # Since capabilities to express dynamic subclasses created by factories
        # like `POINTER` are not introduced into Python's type system yet,
        # it is typed to return `IDictionary`.
        dic0 = comtypes.client.CreateObject(
            Scripting.Dictionary, interface=Scripting.IDictionary
        )
        self.assertIsInstance(dic0, POINTER(Scripting.IDictionary))
        self.assertIsInstance(dic0, Scripting.IDictionary)
        assert_type(dic0, Scripting.IDictionary)
        # `CompareMode` is a normal property.
        self.assertEqual(dic0.CompareMode, Scripting.BinaryCompare)
        dic0.CompareMode = Scripting.TextCompare
        # Ensure that `IDictionary` is callable and subscriptable, and that
        # the `Item` property is a named property.
        dic0["foo"] = 1
        self.assertTrue(dic0["foo"] == dic0.Item["foo"] == 1)
        dic0.Item["bar"] = 2
        self.assertTrue(dic0("bar") == dic0.Item("bar") == 2)
