from pathlib import Path
import sys

from comtypes import IUnknown
from comtypes.automation import IDispatch
from comtypes.client import CreateObject, GetModule


class Test_Static:
	@classmethod
	def setup_class(cls):
		mod = GetModule("scrrun.dll")
		sys.modules.pop(mod.__name__)
		del mod

	class Test_Scripting_Dictionary:
		def test_CreateObjectByCoClass(self):
			from comtypes.gen import Scripting as scrrun

			dic = CreateObject(scrrun.Dictionary)
			assert isinstance(dic, scrrun.IDictionary)
			assert isinstance(dic, IUnknown)
			assert isinstance(dic, IDispatch)

			
