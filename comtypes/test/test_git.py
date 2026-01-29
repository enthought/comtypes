import contextlib
import unittest as ut
from collections.abc import Iterator

from comtypes import IUnknown
from comtypes.git import (
    GetInterfaceFromGlobal,
    RegisterInterfaceInGlobal,
    RevokeInterfaceFromGlobal,
)
from comtypes.typeinfo import CreateTypeLib, ICreateTypeLib


@contextlib.contextmanager
def register_in_global(obj: IUnknown) -> Iterator[int]:
    cookie = RegisterInterfaceInGlobal(obj)
    try:
        yield cookie
    finally:
        RevokeInterfaceFromGlobal(cookie)


class Test(ut.TestCase):
    def test(self):
        tlib = CreateTypeLib("foo.bar")  # we don not save it later
        self.assertEqual((tlib.AddRef(), tlib.Release()), (2, 1))
        with register_in_global(tlib) as cookie:
            # When an object is registered to GIT, `AddRef` is called,
            # incrementing its reference count. This ensures the object
            # remains valid as long as it's globally registered in the GIT.
            self.assertEqual((tlib.AddRef(), tlib.Release()), (3, 2))
            GetInterfaceFromGlobal(cookie, interface=ICreateTypeLib)
            GetInterfaceFromGlobal(cookie, interface=ICreateTypeLib)
            GetInterfaceFromGlobal(cookie)
            self.assertEqual((tlib.AddRef(), tlib.Release()), (3, 2))
        # When an object is revoked from the GIT, `Release` is called,
        # decrementing its reference count. This allows the object to be
        # garbage collected if no other references exist, ensuring proper
        # resource management.
        self.assertEqual((tlib.AddRef(), tlib.Release()), (2, 1))
