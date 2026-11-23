"""Tests for parameterized (named) properties via the native C++ COM test server.

A COM property that takes arguments is exposed by `comtypes` through its
`named_property` descriptor.  The bound form of that descriptor
(`bound_named_property`) supports `obj.Prop[index]`, `obj.Prop[index] = value`,
`obj.Prop(index)` and tuple indexing `obj.Prop[a, b]`.  These tests exercise
that behavior against a real out-of-process COM server, so the argument and
return-value marshalling of parameterized properties is covered as well.

The test double `CoNamedPropertyTest` exposes the same set of parameterized
properties through two interfaces:

* `IDualNamedPropertyTest` - reached through the dual vtable.
* `IDispNamedPropertyTest` - reached through `IDispatch::Invoke`.

Both interfaces are tested because `comtypes` builds the `named_property`
descriptor from two different code paths (`ComPropertyGenerator` for dual
interfaces and `DispPropertyGenerator` for dispinterfaces).

The properties follow this contract:

* Reading an index that has not been assigned returns a VARIANT whose scode is
  DISP_E_BADINDEX, so "not set" can be told apart from a stored value.
"""

import unittest

from comtypes import CLSCTX_LOCAL_SERVER, hresult
from comtypes.client import CreateObject, GetModule

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDualNamedProperty(unittest.TestCase):
    """Parameterized properties reached through the dual vtable interface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.NamedPropertyTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualNamedPropertyTest,
        )

    def test_getitem_unset_index_returns_badindex_variant(self):
        """Reading an index that was never assigned must report DISP_E_BADINDEX.

        The getter returns a VARIANT instead of a plain long so that "not set"
        can be told apart from a stored value.  The VARIANT carries
        DISP_E_BADINDEX in its scode field.  (For the dual vtable path comtypes
        clears the VARIANT to VT_NULL after reading the out parameter, but the
        scode field is preserved, which is the same behavior the VT_ERROR
        round-trip tests rely on.)
        """
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        """`obj.Item[index] = value` must store the value for that index.

        This is the basic read/write contract of a parameterized property:
        the value assigned through `__setitem__` must be returned by a later
        `__getitem__` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_call_passes_index_to_getter(self):
        """`obj.Item(index)` must forward `index` to the property getter.

        `bound_named_property.__call__` is an alternative spelling of
        `__getitem__`; it must reach the same getter with the same argument.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDispNamedProperty(unittest.TestCase):
    """Parameterized properties reached through the dispinterface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.NamedPropertyTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispNamedPropertyTest,
        )

    def test_getitem_unset_index_returns_badindex_variant(self):
        """Reading an index that was never assigned must report DISP_E_BADINDEX.

        The getter returns a VARIANT instead of a plain long so that "not set"
        can be told apart from a stored value.  The VARIANT carries
        DISP_E_BADINDEX in its scode field.
        """
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        """`obj.Item[index] = value` must store the value for that index.

        This is the basic read/write contract of a parameterized property:
        the value assigned through `__setitem__` must be returned by a later
        `__getitem__` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_call_passes_index_to_getter(self):
        """`obj.Item(index)` must forward `index` to the property getter.

        `bound_named_property.__call__` is an alternative spelling of
        `__getitem__`; it must reach the same getter with the same argument.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)


if __name__ == "__main__":
    unittest.main()
