import unittest as ut

from comtypes.git import (
    GetInterfaceFromGlobal,
    RegisterInterfaceInGlobal,
    RevokeInterfaceFromGlobal,
)
from comtypes.typeinfo import CreateTypeLib, ICreateTypeLib


class Test(ut.TestCase):
    def test(self):
        tlib = CreateTypeLib("foo.bar")  # we don not save it later
        self.assertEqual((tlib.AddRef(), tlib.Release()), (2, 1))
        cookie = RegisterInterfaceInGlobal(tlib)
        # When an object is registered to GIT, `AddRef` is called, incrementing
        # its reference count. This ensures the object remains valid as long as
        # it's globally registered in the GIT.
        self.assertEqual((tlib.AddRef(), tlib.Release()), (3, 2))
        GetInterfaceFromGlobal(cookie, interface=ICreateTypeLib)
        GetInterfaceFromGlobal(cookie, interface=ICreateTypeLib)
        GetInterfaceFromGlobal(cookie)
        self.assertEqual((tlib.AddRef(), tlib.Release()), (3, 2))
        RevokeInterfaceFromGlobal(cookie)
        # When an object is revoked from the GIT, `Release` is called,
        # decrementing its reference count. This allows the object to be
        # garbage collected if no other references exist, ensuring proper
        # resource management.
        self.assertEqual((tlib.AddRef(), tlib.Release()), (2, 1))
