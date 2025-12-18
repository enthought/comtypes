"""Use Scripting.Dictionary to test the lazybind module."""

import unittest

from comtypes import typeinfo
from comtypes.automation import VARIANT
from comtypes.client import CreateObject, GetModule
from comtypes.client.lazybind import Dispatch

GetModule("scrrun.dll")
import comtypes.gen.Scripting as scrrun  # noqa


class Test(unittest.TestCase):
    def test_dynamic(self):
        d = CreateObject("Scripting.Dictionary", dynamic=True)
        self.assertEqual(type(d), Dispatch)

        # Count is a normal propget, no propput
        self.assertEqual(d.Count, 0)
        with self.assertRaises(AttributeError):
            setattr(d, "Count", -1)

        # HashVal is a 'named' propget, no propput
        # HashVal is a 'hidden' member and used internally.
        ##d.HashVal

        # Add(Key, Item) -> None
        self.assertEqual(d.Add("one", 1), None)
        self.assertEqual(d.Count, 1)

        # RemoveAll() -> None
        self.assertEqual(d.RemoveAll(), None)
        self.assertEqual(d.Count, 0)

        # CompareMode: propget, propput
        # (Can only be set when dict is empty!)
        # Verify that the default is BinaryCompare.
        self.assertEqual(d.CompareMode, scrrun.BinaryCompare)
        d.CompareMode = scrrun.TextCompare
        self.assertEqual(d.CompareMode, scrrun.TextCompare)
        d.CompareMode = scrrun.BinaryCompare

        # Exists(key) -> bool
        self.assertEqual(d.Exists(42), False)
        d.Add(42, "foo")
        self.assertEqual(d.Exists(42), True)

        # Keys() -> array
        # Items() -> array
        self.assertEqual(d.Keys(), (42,))
        self.assertEqual(d.Items(), ("foo",))
        d.Remove(42)
        self.assertEqual(d.Exists(42), False)
        self.assertEqual(d.Keys(), ())
        self.assertEqual(d.Items(), ())

        # Item[key] : propget
        d.Add(42, "foo")
        self.assertEqual(d.Item[42], "foo")

        d.Add("spam", "bar")
        self.assertEqual(d.Item["spam"], "bar")

        # Item[key] = value: propput, propputref
        d.Item["key"] = "value"
        self.assertEqual(d.Item["key"], "value")
        d.Item[42] = 73, 48
        self.assertEqual(d.Item[42], (73, 48))

        ################################################################
        # part 2, testing propput and propputref

        s = CreateObject("Scripting.Dictionary", dynamic=True)
        s.CompareMode = scrrun.DatabaseCompare

        # This calls propputref, since we assign an Object
        d.Item["object"] = s
        # This calls propput, since we assign a Value
        d.Item["value"] = s.CompareMode

        a = d.Item["object"]

        self.assertEqual(d.Item["object"], s)
        self.assertEqual(d.Item["object"].CompareMode, scrrun.DatabaseCompare)
        self.assertEqual(d.Item["value"], scrrun.DatabaseCompare)

        # Changing a property of the object
        s.CompareMode = scrrun.BinaryCompare
        self.assertEqual(d.Item["object"], s)
        self.assertEqual(d.Item["object"].CompareMode, scrrun.BinaryCompare)
        self.assertEqual(d.Item["value"], scrrun.DatabaseCompare)

        # This also calls propputref since we assign an Object
        d.Item["var"] = VARIANT(s)
        self.assertEqual(d.Item["var"], s)

        # iter(d)
        keys = [x for x in d]
        self.assertEqual(d.Keys(), tuple([x for x in d]))

        # d[key] = value
        # d[key] -> value
        d["blah"] = "blarg"
        self.assertEqual(d["blah"], "blarg")
        # d(key) -> value
        self.assertEqual(d("blah"), "blarg")

    def test_static(self):
        d = CreateObject(scrrun.Dictionary, interface=scrrun.IDictionary)
        # This confirms that the Dictionary is a dual interface.
        self.assertTrue(
            d.GetTypeInfo(0).GetTypeAttr().wTypeFlags & typeinfo.TYPEFLAG_FDUAL
        )
        # Dual interfaces call COM methods that support named arguments.
        d.Add("one", 1)
        d.Add("two", Item=2)
        d.Add(Key="three", Item=3)
        d.Add(Item=4, Key="four")
        d.Item["five"] = 5
        d["six"] = 6
        self.assertEqual(d.Count, 6)
        self.assertEqual(len(d), 6)
        self.assertEqual(d("six"), 6)
        self.assertEqual(d.Item("five"), 5)
        self.assertEqual(d("four"), 4)
        self.assertEqual(d["three"], 3)
        self.assertEqual(d.Item["two"], 2)
        self.assertEqual(d("one"), 1)
        # NOTE: Named parameters are not yet implemented for the named property.
        # See https://github.com/enthought/comtypes/issues/371
        # TODO: After named parameters are supported, this will become a test to
        # assert the return value.
        with self.assertRaises(TypeError):
            d.Item(Key="two")
        with self.assertRaises(TypeError):
            d(Key="one")


if __name__ == "__main__":
    unittest.main()
