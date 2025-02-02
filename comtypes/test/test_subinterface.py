import sys
import unittest
from ctypes import *

from comtypes import GUID, IUnknown


def test_main():
    from test import test_support

    test_support.run_unittest(Test)


class Test(unittest.TestCase):
    def test_subinterface(self):
        class ISub(IUnknown):
            pass

    def test_subclass(self):
        class X(c_void_p):
            pass
