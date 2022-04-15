from pathlib import Path
import sys

from comtypes import IUnknown
from comtypes.automation import IDispatch
from comtypes.client.lazybind import Dispatch
from comtypes.client import CreateObject, GetModule

import pytest


@pytest.fixture(scope="module", autouse=True)
def _setup():
	mod = GetModule("scrrun.dll")
	sys.modules.pop(mod.__name__)
	del mod


class Test_Scripting_Dictionary:
	def test_CreateObject_TakesCoClass(self):
		from comtypes.gen import Scripting as scrrun

		dic = CreateObject(scrrun.Dictionary)
		assert isinstance(dic, scrrun.IDictionary)
		assert isinstance(dic, IUnknown)
		assert isinstance(dic, IDispatch)
		assert len(dic) == 0
		dic["foo"] = 1
		assert dic.Count == 1
		dic["bar"] = 3.14
		dic["baz"] = "qux"
		assert len(dic) == 3
		assert dic("foo") == 1
		assert dic("bar") == 3.14
		assert dic["baz"] == "qux"
		dic.RemoveAll()
		assert dic.Count == 0
		dic.Add("abc", 2)
		assert dic.Exists("abc") is True
		assert dic.Exists("lmn") is False
		dic.Add("lmn", 1.414)
		assert len(dic) == 2
		dic.Add("xyz", "uvw")
		assert dic.Count == 3
		assert [k for k in dic.Keys()] == [k for k in dic]

	def test_CreateObject_TakesClsIDAndDynamicTrue(self):
		dic = CreateObject("Scripting.Dictionary", dynamic=True)
		assert isinstance(dic, Dispatch)
		dic["foo"] = 1
		assert dic.Count == 1
		dic["bar"] = 3.14
		dic["baz"] = "qux"
		assert dic("foo") == 1
		assert dic("bar") == 3.14
		assert dic["baz"] == "qux"
		dic.RemoveAll()
		assert dic.Count == 0
		dic.Add("abc", 2)
		assert dic.Exists("abc") is True
		assert dic.Exists("lmn") is False
		dic.Add("lmn", 1.414)
		dic.Add("xyz", "uvw")
		assert dic.Count == 3
		assert [k for k in dic.Keys()] == [k for k in dic]
