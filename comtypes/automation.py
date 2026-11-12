# comtypes.automation module
from _ctypes import COMError
from collections.abc import Sequence
from ctypes import HRESULT, POINTER, byref, c_int, c_uint, c_ulong, c_wchar_p, pointer
from ctypes.wintypes import UINT, WORD
from ctypes.wintypes import VARIANT_BOOL as VARIANT_BOOL
from ctypes.wintypes import WCHAR as WCHAR
from typing import TYPE_CHECKING, Any, ClassVar, Literal, overload

import comtypes
import comtypes.typeinfo
from comtypes import BSTR as BSTR
from comtypes import COMMETHOD, GUID, IID, STDMETHOD, IUnknown
from comtypes import hresult as hresult
from comtypes._memberspec import DISPATCH_METHOD as DISPATCH_METHOD
from comtypes._memberspec import DISPATCH_PROPERTYGET as DISPATCH_PROPERTYGET
from comtypes._memberspec import DISPATCH_PROPERTYPUT as DISPATCH_PROPERTYPUT
from comtypes._memberspec import DISPATCH_PROPERTYPUTREF as DISPATCH_PROPERTYPUTREF
from comtypes.safearray import _midlSAFEARRAY as _midlSAFEARRAY

if TYPE_CHECKING:
    from ctypes import _CArgObject, _CDataType

    from comtypes import hints  # type: ignore
    from comtypes._memberspec import _DispMemberSpec
else:
    _CArgObject = type(byref(c_int()))

# isort: off
from comtypes._automation import LCID as LCID
from comtypes._automation import DISPID as DISPID
from comtypes._automation import SCODE as SCODE

from comtypes._automation import VARTYPE as VARTYPE

from comtypes._automation import tagINVOKEKIND as tagINVOKEKIND

# isort: on
INVOKE_FUNC = DISPATCH_METHOD
INVOKE_PROPERTYGET = DISPATCH_PROPERTYGET
INVOKE_PROPERTYPUT = DISPATCH_PROPERTYPUT
INVOKE_PROPERTYPUTREF = DISPATCH_PROPERTYPUTREF
INVOKEKIND = tagINVOKEKIND


################################
# helpers
# isort: off
from comtypes._automation import IID_NULL as IID_NULL
from comtypes._automation import riid_null as riid_null
from comtypes._automation import _com_null_date as _com_null_date
# isort: on

################################################################
# VARIANT, in all it's glory.

# isort: off
from comtypes._automation import VARENUM as VARENUM  # enum
from comtypes._automation import VT_EMPTY as VT_EMPTY
from comtypes._automation import VT_NULL as VT_NULL
from comtypes._automation import VT_I2 as VT_I2
from comtypes._automation import VT_I4 as VT_I4
from comtypes._automation import VT_R4 as VT_R4
from comtypes._automation import VT_R8 as VT_R8
from comtypes._automation import VT_CY as VT_CY
from comtypes._automation import VT_DATE as VT_DATE
from comtypes._automation import VT_BSTR as VT_BSTR
from comtypes._automation import VT_DISPATCH as VT_DISPATCH
from comtypes._automation import VT_ERROR as VT_ERROR
from comtypes._automation import VT_BOOL as VT_BOOL
from comtypes._automation import VT_VARIANT as VT_VARIANT
from comtypes._automation import VT_UNKNOWN as VT_UNKNOWN
from comtypes._automation import VT_DECIMAL as VT_DECIMAL
from comtypes._automation import VT_I1 as VT_I1
from comtypes._automation import VT_UI1 as VT_UI1
from comtypes._automation import VT_UI2 as VT_UI2
from comtypes._automation import VT_UI4 as VT_UI4
from comtypes._automation import VT_I8 as VT_I8
from comtypes._automation import VT_UI8 as VT_UI8
from comtypes._automation import VT_INT as VT_INT
from comtypes._automation import VT_UINT as VT_UINT
from comtypes._automation import VT_VOID as VT_VOID
from comtypes._automation import VT_HRESULT as VT_HRESULT
from comtypes._automation import VT_PTR as VT_PTR
from comtypes._automation import VT_SAFEARRAY as VT_SAFEARRAY
from comtypes._automation import VT_CARRAY as VT_CARRAY
from comtypes._automation import VT_USERDEFINED as VT_USERDEFINED
from comtypes._automation import VT_LPSTR as VT_LPSTR
from comtypes._automation import VT_LPWSTR as VT_LPWSTR
from comtypes._automation import VT_RECORD as VT_RECORD
from comtypes._automation import VT_INT_PTR as VT_INT_PTR
from comtypes._automation import VT_UINT_PTR as VT_UINT_PTR
from comtypes._automation import VT_FILETIME as VT_FILETIME
from comtypes._automation import VT_BLOB as VT_BLOB
from comtypes._automation import VT_STREAM as VT_STREAM
from comtypes._automation import VT_STORAGE as VT_STORAGE
from comtypes._automation import VT_STREAMED_OBJECT as VT_STREAMED_OBJECT
from comtypes._automation import VT_STORED_OBJECT as VT_STORED_OBJECT
from comtypes._automation import VT_BLOB_OBJECT as VT_BLOB_OBJECT
from comtypes._automation import VT_CF as VT_CF
from comtypes._automation import VT_CLSID as VT_CLSID
from comtypes._automation import VT_VERSIONED_STREAM as VT_VERSIONED_STREAM
from comtypes._automation import VT_BSTR_BLOB as VT_BSTR_BLOB
from comtypes._automation import VT_VECTOR as VT_VECTOR
from comtypes._automation import VT_ARRAY as VT_ARRAY
from comtypes._automation import VT_BYREF as VT_BYREF
from comtypes._automation import VT_RESERVED as VT_RESERVED
from comtypes._automation import VT_ILLEGAL as VT_ILLEGAL
from comtypes._automation import VT_ILLEGALMASKED as VT_ILLEGALMASKED
from comtypes._automation import VT_TYPEMASK as VT_TYPEMASK

# isort: on
from comtypes._automation import tagCY as tagCY

CY = tagCY
CURRENCY = CY
from comtypes._automation import tagDEC as tagDEC

DECIMAL = tagDEC

from comtypes._automation import tagVARIANT as tagVARIANT

VARIANT = tagVARIANT
VARIANTARG = VARIANT

from comtypes._automation import _SysAllocStringLen as _SysAllocStringLen
from comtypes._automation import _VariantChangeType as _VariantChangeType
from comtypes._automation import _VariantClear as _VariantClear
from comtypes._automation import _VariantCopy as _VariantCopy
from comtypes._automation import _VariantCopyInd as _VariantCopyInd


################################################################
# interfaces, structures, ...
class IEnumVARIANT(IUnknown):
    _iid_ = GUID("{00020404-0000-0000-C000-000000000046}")
    _idlflags_ = ["hidden"]
    _dynamic = False

    if TYPE_CHECKING:

        def Skip(self, cConnections: int) -> hints.Hresult: ...
        def Reset(self) -> hints.Hresult: ...
        def Clone(self) -> hints.Self: ...

    def __iter__(self) -> "hints.Self":
        return self

    def __next__(self) -> Any:
        item, fetched = self.Next(1)
        if fetched:
            return item
        raise StopIteration

    def __getitem__(self, index: int) -> Any:
        self.Reset()
        # Does not yet work.
        # if isinstance(index, slice):
        #     self.Skip(index.start or 0)
        #     return self.Next(index.stop or sys.maxint)
        self.Skip(index)
        item, fetched = self.Next(1)
        if fetched:
            return item
        raise IndexError

    @overload
    def Next(self, celt: Literal[1]) -> tuple[Any, int]: ...
    @overload
    def Next(self, celt: int) -> Sequence[Any]: ...
    def Next(self, celt):  # type: ignore
        """Retrieve the next *celt* items from the enumeration.

        This method behaves differently depending on the value of *celt*:

        - `celt == 1`:
          A single `VARIANT` is fetched via one COM call.  The return
          value is `(value, fetched)` — a two-element tuple where *value*
          is the retrieved object and *fetched* is the number of items
          actually returned (`0` or `1`).

        - `celt != 1` (including `0`):
          A `VARIANT` array of length *celt* is allocated and filled in a
          single COM call.  Only the first *fetched* slots are meaningful;
          the rest are discarded.  The return value is fetched items
          (possibly empty when `celt == 0` or nothing is left in the
          enumeration).

        Args:
            celt: The maximum number of items to retrieve.

        Returns:
            A `(value, fetched)` tuple when *celt* is `1`.
            Fetched items when *celt* is not `1`.

        Note:
            This object implements dunder methods that define iterator and
            container behavior, so a more Pythonic approach is recommended
            for accessing its elements rather than calling this method
            directly.
        """
        # This wrapper deviates from a plain COM `IEnumVARIANT::Next` proxy in
        # two ways that are specific to this package:
        #
        # 1. Support for celt != 1 (commit 9f68b6a, by theller):
        #    "this allows to get more objects at a time."
        #    When celt != 1, the method allocates a VARIANT array, fetches up
        #    to `celt` items in a single COM call, and returns them as a list.
        #
        # 2. Return type for celt == 1 (commit 65bdc13, by theller):
        #    The original override returned only the unwrapped value. It was
        #    later corrected so that celt == 1 returns a tuple corresponding
        #    exactly to the two `[out]` parameters declared by the COM method
        #    specifier.
        #
        # Both decisions are those of the package originator (theller) and are
        # intentionally preserved here.
        fetched = c_ulong()
        if celt == 1:
            v = VARIANT()
            self.__com_Next(celt, v, fetched)  # type: ignore
            return v._get_value(dynamic=self._dynamic), fetched.value
        array = (VARIANT * celt)()
        self.__com_Next(celt, array, fetched)  # type: ignore
        result = [v._get_value(dynamic=self._dynamic) for v in array[: fetched.value]]
        # Release VARIANT refcounts before the temporary array is freed.
        for v in array:
            v.value = None
        return result


IEnumVARIANT._methods_ = [
    COMMETHOD(
        [],
        HRESULT,
        "Next",
        (["in"], c_ulong, "celt"),
        (["out"], POINTER(VARIANT), "rgvar"),
        (["out"], POINTER(c_ulong), "pceltFetched"),
    ),
    COMMETHOD([], HRESULT, "Skip", (["in"], c_ulong, "celt")),
    COMMETHOD([], HRESULT, "Reset"),
    COMMETHOD(
        [], HRESULT, "Clone", (["out"], POINTER(POINTER(IEnumVARIANT)), "ppenum")
    ),
]

from comtypes._automation import tagEXCEPINFO as tagEXCEPINFO

EXCEPINFO = tagEXCEPINFO

from comtypes._automation import tagDISPPARAMS as tagDISPPARAMS

DISPPARAMS = tagDISPPARAMS


# isort: off
from comtypes._automation import DISPID_VALUE as DISPID_VALUE
from comtypes._automation import DISPID_UNKNOWN as DISPID_UNKNOWN
from comtypes._automation import DISPID_PROPERTYPUT as DISPID_PROPERTYPUT
from comtypes._automation import DISPID_NEWENUM as DISPID_NEWENUM
from comtypes._automation import DISPID_EVALUATE as DISPID_EVALUATE
from comtypes._automation import DISPID_CONSTRUCTOR as DISPID_CONSTRUCTOR
from comtypes._automation import DISPID_DESTRUCTOR as DISPID_DESTRUCTOR
from comtypes._automation import DISPID_COLLECT as DISPID_COLLECT
# isort: on


class IDispatch(IUnknown):
    _disp_methods_: ClassVar[list["_DispMemberSpec"]]

    _iid_ = GUID("{00020400-0000-0000-C000-000000000046}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetTypeInfoCount", (["out"], POINTER(UINT))),
        COMMETHOD(
            [],
            HRESULT,
            "GetTypeInfo",
            (["in"], UINT, "index"),
            (["in"], LCID, "lcid", 0),
            # Normally, we would declare this parameter in this way:
            # (['out'], POINTER(POINTER(ITypeInfo)) ) ),
            # this is a remnant of when import comtypes.typeinfo at the top
            # level used to cause recursive imports.
            (["out"], POINTER(POINTER(IUnknown))),
        ),
        STDMETHOD(
            HRESULT,
            "GetIDsOfNames",
            [POINTER(IID), POINTER(c_wchar_p), UINT, LCID, POINTER(DISPID)],
        ),
        STDMETHOD(
            HRESULT,
            "Invoke",
            [
                DISPID,
                POINTER(IID),
                LCID,
                WORD,
                POINTER(DISPPARAMS),
                POINTER(VARIANT),
                POINTER(EXCEPINFO),
                POINTER(UINT),
            ],
        ),
    ]

    if TYPE_CHECKING:

        def GetTypeInfoCount(self) -> int: ...

    def GetTypeInfo(self, index: int, lcid: int = 0) -> "hints.ITypeInfo":
        """Return type information.  Index 0 specifies typeinfo for IDispatch"""
        result = self._GetTypeInfo(index, lcid)  # type: ignore
        return result.QueryInterface(comtypes.typeinfo.ITypeInfo)

    def GetIDsOfNames(self, *names: str, **kw: Any) -> list[int]:
        """Map string names to integer ids."""
        lcid = kw.pop("lcid", 0)
        assert not kw
        arr = (c_wchar_p * len(names))(*names)
        ids = (DISPID * len(names))()
        self.__com_GetIDsOfNames(riid_null, arr, len(names), lcid, ids)  # type: ignore
        return ids[:]

    def _invoke(self, memid: int, invkind: int, lcid: int, *args: Any) -> Any:
        var = VARIANT()
        argerr = c_uint()
        dp = DISPPARAMS()

        if args:
            array = (VARIANT * len(args))()

            for i, a in enumerate(args[::-1]):
                array[i].value = a

            dp.cArgs = len(args)
            if invkind in (DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF):
                dp.cNamedArgs = 1
                dp.rgdispidNamedArgs = pointer(DISPID(DISPID_PROPERTYPUT))
            dp.rgvarg = array

        self.__com_Invoke(  # type: ignore
            memid, riid_null, lcid, invkind, dp, var, None, argerr
        )
        return var._get_value(dynamic=True)

    def __make_dp(self, _invkind: int, *args: Any) -> DISPPARAMS:
        array = (VARIANT * len(args))()
        for i, a in enumerate(args[::-1]):
            array[i].value = a
        dp = DISPPARAMS()
        dp.cArgs = len(args)
        dp.rgvarg = array
        if _invkind in (DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF):  # propput
            dp.cNamedArgs = 1
            dp.rgdispidNamedArgs = pointer(DISPID(DISPID_PROPERTYPUT))
        else:
            dp.cNamedArgs = 0
        return dp

    def Invoke(self, dispid: int, *args: Any, **kw: Any) -> Any:
        """Invoke a method or property."""

        # Memory management in Dispatch::Invoke calls:
        # http://msdn.microsoft.com/library/en-us/automat/htm/chap5_4x2q.asp
        # Quote:
        #     The *CALLING* code is responsible for releasing all strings and
        #     objects referred to by rgvarg[ ] or placed in *pVarResult.
        #
        # For comtypes this is handled in DISPPARAMS.__del__ and VARIANT.__del__.
        _invkind = kw.pop("_invkind", DISPATCH_METHOD)
        _lcid = kw.pop("_lcid", 0)
        if kw:
            raise ValueError("named parameters not yet implemented")
        dp = self.__make_dp(_invkind, *args)
        result = VARIANT()
        excepinfo = EXCEPINFO()
        argerr = c_uint()
        try:
            self.__com_Invoke(  # type: ignore
                dispid,
                riid_null,
                _lcid,
                _invkind,
                byref(dp),
                byref(result),
                byref(excepinfo),
                byref(argerr),
            )
        except COMError as err:
            (hr, text, details) = err.args
            if hr == hresult.DISP_E_EXCEPTION:
                details = (
                    excepinfo.bstrDescription,
                    excepinfo.bstrSource,
                    excepinfo.bstrHelpFile,
                    excepinfo.dwHelpContext,
                    excepinfo.scode,
                )
                raise COMError(hr, text, details)
            elif hr == hresult.DISP_E_PARAMNOTFOUND:
                # MSDN says: You get the error DISP_E_PARAMNOTFOUND
                # when you try to set a property and you have not
                # initialized the cNamedArgs and rgdispidNamedArgs
                # elements of your DISPPARAMS structure.
                #
                # So, this looks like a bug.
                raise COMError(hr, text, argerr.value)
            elif hr == hresult.DISP_E_TYPEMISMATCH:
                # MSDN: One or more of the arguments could not be
                # coerced.
                #
                # Hm, should we raise TypeError, or COMError?
                raise COMError(
                    hr, text, (f"TypeError: Parameter {argerr.value + 1}", args)
                )
            raise
        return result._get_value(dynamic=True)

    # XXX Would separate methods for _METHOD, _PROPERTYGET and _PROPERTYPUT be better?


################################################################
# safearrays

# map ctypes types to VARTYPE values
from comtypes._automation import _arraycode_to_vartype as _arraycode_to_vartype
from comtypes._automation import _ctype_to_vartype as _ctype_to_vartype
from comtypes._automation import _vartype_to_ctype as _vartype_to_ctype

# fmt: off
__known_symbols__ = [
    "CURRENCY", "CY", "tagCY", "DECIMAL", "tagDEC", "DISPATCH_METHOD",
    "DISPATCH_PROPERTYGET", "DISPATCH_PROPERTYPUT", "DISPATCH_PROPERTYPUTREF",
    "DISPID", "DISPID_COLLECT", "DISPID_CONSTRUCTOR", "DISPID_DESTRUCTOR",
    "DISPID_EVALUATE", "DISPID_NEWENUM", "DISPID_PROPERTYPUT",
    "DISPID_UNKNOWN", "DISPID_VALUE", "DISPPARAMS", "tagDISPPARAMS",
    "EXCEPINFO", "tagEXCEPINFO", "IDispatch", "IEnumVARIANT", "IID_NULL",
    "INVOKE_FUNC", "INVOKE_PROPERTYGET", "INVOKE_PROPERTYPUT",
    "INVOKE_PROPERTYPUTREF", "INVOKEKIND", "tagINVOKEKIND", "_midlSAFEARRAY",
    "SCODE", "_SysAllocStringLen", "VARENUM", "VARIANT", "tagVARIANT",
    "VARIANTARG", "_VariantChangeType", "_VariantClear", "_VariantCopy",
    "_VariantCopyInd", "VARTYPE", "VT_ARRAY", "VT_BLOB", "VT_BLOB_OBJECT",
    "VT_BOOL", "VT_BSTR", "VT_BSTR_BLOB", "VT_BYREF", "VT_CARRAY", "VT_CF",
    "VT_CLSID", "VT_CY", "VT_DATE", "VT_DECIMAL", "VT_DISPATCH", "VT_EMPTY",
    "VT_ERROR", "VT_FILETIME", "VT_HRESULT", "VT_I1", "VT_I2", "VT_I4",
    "VT_I8", "VT_ILLEGAL", "VT_ILLEGALMASKED", "VT_INT", "VT_INT_PTR",
    "VT_LPSTR", "VT_LPWSTR", "VT_NULL", "VT_PTR", "VT_R4", "VT_R8",
    "VT_RECORD", "VT_RESERVED", "VT_SAFEARRAY", "VT_STORAGE",
    "VT_STORED_OBJECT", "VT_STREAM", "VT_STREAMED_OBJECT", "VT_TYPEMASK",
    "VT_UI1", "VT_UI2", "VT_UI4", "VT_UI8", "VT_UINT", "VT_UINT_PTR",
    "VT_UNKNOWN", "VT_USERDEFINED", "VT_VARIANT", "VT_VECTOR",
    "VT_VERSIONED_STREAM", "VT_VOID",
]
# fmt: on
