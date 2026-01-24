import unittest

import comtypes.client
from comtypes.automation import IEnumVARIANT
from comtypes.server.automation import VARIANTEnumerator

comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting as scrrun


class TestVARIANTEnumerator(unittest.TestCase):
    def setUp(self):
        # Create a list of IDispatch objects to enumerate
        dict1 = comtypes.client.CreateObject(
            "Scripting.Dictionary", interface=scrrun.IDictionary
        )
        dict1.Add("key1", "value1")
        dict2 = comtypes.client.CreateObject(
            "Scripting.Dictionary", interface=scrrun.IDictionary
        )
        dict2.Add("key2", "value2")
        dict3 = comtypes.client.CreateObject(
            "Scripting.Dictionary", interface=scrrun.IDictionary
        )
        dict3.Add("key3", "value3")
        self.items = [dict1, dict2, dict3]
        self.enumerator = VARIANTEnumerator(self.items)

    def test_Next(self):
        enum_variant = self.enumerator.QueryInterface(IEnumVARIANT)
        # Retrieve the first item
        item, fetched = enum_variant.Next(1)
        self.assertEqual(fetched, 1)
        dict1 = item.QueryInterface(scrrun.IDictionary)
        self.assertEqual(dict1.Item("key1"), "value1")
        # Retrieve the second item
        item, fetched = enum_variant.Next(1)
        dict2 = item.QueryInterface(scrrun.IDictionary)
        self.assertEqual(dict2.Item("key2"), "value2")
        # Retrieve the third item
        item, fetched = enum_variant.Next(1)
        self.assertEqual(fetched, 1)
        dict3 = item.QueryInterface(scrrun.IDictionary)
        self.assertEqual(dict3.Item("key3"), "value3")
        # After all items are enumerated, `Next` should return 0 fetched
        item, fetched = enum_variant.Next(1)
        self.assertEqual(fetched, 0)
        self.assertFalse(item)
