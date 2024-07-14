# comtypes.automation module
import array
import datetime
import decimal
import sys
from ctypes import *
from ctypes import _Pointer
from _ctypes import CopyComPointer
from ctypes.wintypes import DWORD, LONG, UINT, VARIANT_BOOL, WCHAR, WORD
from typing import Any, ClassVar, overload, TYPE_CHECKING
from typing import Optional, Union as _UnionT
from typing import Dict, List, Tuple, Type
from typing import Callable, Sequence

from comtypes import _CData, BSTR, COMError, COMMETHOD, GUID, IID, IUnknown, STDMETHOD
from comtypes.hresult import *
from comtypes._memberspec import _DispMemberSpec
import comtypes.patcher
import comtypes

if TYPE_CHECKING:
    from comtypes import hints  # type: ignore
    from comtypes import _safearray
else:
    try:
        from comtypes import _safearray
    except (ImportError, AttributeError):

        class _safearray(object):
            tagSAFEARRAY = None


LCID = DWORD
DISPID = LONG
SCODE = LONG

VARTYPE = c_ushort

DISPATCH_METHOD = 1
DISPATCH_PROPERTYGET = 2
DISPATCH_PROPERTYPUT = 4
DISPATCH_PROPERTYPUTREF = 8

tagINVOKEKIND = c_int
INVOKE_FUNC = DISPATCH_METHOD
INVOKE_PROPERTYGET = DISPATCH_PROPERTYGET
INVOKE_PROPERTYPUT = DISPATCH_PROPERTYPUT
INVOKE_PROPERTYPUTREF = DISPATCH_PROPERTYPUTREF
INVOKEKIND = tagINVOKEKIND


################################
# helpers
IID_NULL = GUID()
riid_null = byref(IID_NULL)
_byref_type = type(byref(c_int()))

# 30. December 1899, midnight.  For VT_DATE.
_com_null_date = datetime.datetime(1899, 12, 30, 0, 0, 0)

################################################################
# VARIANT, in all it's glory.
VARENUM = c_int  # enum
VT_EMPTY = 0
VT_NULL = 1
VT_I2 = 2
VT_I4 = 3
VT_R4 = 4
VT_R8 = 5
VT_CY = 6
VT_DATE = 7
VT_BSTR = 8
VT_DISPATCH = 9
VT_ERROR = 10
VT_BOOL = 11
VT_VARIANT = 12
VT_UNKNOWN = 13
VT_DECIMAL = 14
VT_I1 = 16
VT_UI1 = 17
VT_UI2 = 18
VT_UI4 = 19
VT_I8 = 20
VT_UI8 = 21
VT_INT = 22
VT_UINT = 23
VT_VOID = 24
VT_HRESULT = 25
VT_PTR = 26
VT_SAFEARRAY = 27
VT_CARRAY = 28
VT_USERDEFINED = 29
VT_LPSTR = 30
VT_LPWSTR = 31
VT_RECORD = 36
VT_INT_PTR = 37
VT_UINT_PTR = 38
VT_FILETIME = 64
VT_BLOB = 65
VT_STREAM = 66
VT_STORAGE = 67
VT_STREAMED_OBJECT = 68
VT_STORED_OBJECT = 69
VT_BLOB_OBJECT = 70
VT_CF = 71
VT_CLSID = 72
VT_VERSIONED_STREAM = 73
VT_BSTR_BLOB = 4095
VT_VECTOR = 4096
VT_ARRAY = 8192
VT_BYREF = 16384
VT_RESERVED = 32768
VT_ILLEGAL = 65535
VT_ILLEGALMASKED = 4095
VT_TYPEMASK = 4095


class tagCY(Structure):
    _fields_ = [("int64", c_longlong)]


CY = tagCY
CURRENCY = CY


class tagDEC(Structure):
    _fields_ = [
        ("wReserved", c_ushort),
        ("scale", c_ubyte),
        ("sign", c_ubyte),
        ("Hi32", c_ulong),
        ("Lo64", c_ulonglong),
    ]

    def as_decimal(self):
        """Convert a tagDEC struct to Decimal.

        See http://msdn.microsoft.com/en-us/library/cc234586.aspx for the tagDEC
        specification.

        """
        digits = (self.Hi32 << 64) + self.Lo64
        decimal_str = "{0}{1}e-{2}".format(
            "-" if self.sign else "",
            digits,
            self.scale,
        )
        return decimal.Decimal(decimal_str)


DECIMAL = tagDEC


# The VARIANT structure is a good candidate for implementation in a C
# helper extension.  At least the get/set methods.
class tagVARIANT(Structure):
    if TYPE_CHECKING:
        vt: int
        _: "U_VARIANT1.__tagVARIANT.U_VARIANT2"
        null: ClassVar["VARIANT"]
        empty: ClassVar["VARIANT"]
        missing: ClassVar["VARIANT"]

    class U_VARIANT1(Union):
        class __tagVARIANT(Structure):
            # The C Header file defn of VARIANT is much more complicated, but
            # this is the ctypes version - functional as well.
            class U_VARIANT2(Union):
                class _tagBRECORD(Structure):
                    _fields_ = [("pvRecord", c_void_p), ("pRecInfo", POINTER(IUnknown))]

                _fields_ = [
                    ("VT_BOOL", VARIANT_BOOL),
                    ("VT_I1", c_byte),
                    ("VT_I2", c_short),
                    ("VT_I4", c_long),
                    ("VT_I8", c_longlong),
                    ("VT_INT", c_int),
                    ("VT_UI1", c_ubyte),
                    ("VT_UI2", c_ushort),
                    ("VT_UI4", c_ulong),
                    ("VT_UI8", c_ulonglong),
                    ("VT_UINT", c_uint),
                    ("VT_R4", c_float),
                    ("VT_R8", c_double),
                    ("VT_CY", c_longlong),
                    ("c_wchar_p", c_wchar_p),
                    ("c_void_p", c_void_p),
                    ("pparray", POINTER(POINTER(_safearray.tagSAFEARRAY))),
                    ("bstrVal", BSTR),
                    ("_tagBRECORD", _tagBRECORD),
                ]
                _anonymous_ = ["_tagBRECORD"]

            _fields_ = [
                ("vt", VARTYPE),
                ("wReserved1", c_ushort),
                ("wReserved2", c_ushort),
                ("wReserved3", c_ushort),
                ("_", U_VARIANT2),
            ]

        _fields_ = [("__VARIANT_NAME_2", __tagVARIANT), ("decVal", DECIMAL)]
        _anonymous_ = ["__VARIANT_NAME_2"]

    _fields_ = [("__VARIANT_NAME_1", U_VARIANT1)]
    _anonymous_ = ["__VARIANT_NAME_1"]

    def __init__(self, *args):
        if args:
            self.value = args[0]

    def __del__(self):
        if self._b_needsfree_:
            # XXX This does not work.  _b_needsfree_ is never
            # set because the buffer is internal to the object.
            _VariantClear(self)

    def __repr__(self):
        if self.vt & VT_BYREF:
            return "VARIANT(vt=0x%x, byref(%r))" % (self.vt, self[0])
        elif self is type(self).null:
            return "VARIANT.null"
        elif self is type(self).empty:
            return "VARIANT.empty"
        elif self is type(self).missing:
            return "VARIANT.missing"
        return "VARIANT(vt=0x%x, %r)" % (self.vt, self.value)

    @classmethod
    def from_param(cls, value):
        if isinstance(value, cls):
            return value
        return cls(value)

    def __setitem__(self, index, value):
        # This method allows to change the value of a
        # (VT_BYREF|VT_xxx) variant in place.
        if index != 0:
            raise IndexError(index)
        if not self.vt & VT_BYREF:
            raise TypeError("set_byref requires a VT_BYREF VARIANT instance")
        typ = _vartype_to_ctype[self.vt & ~VT_BYREF]
        cast(self._.c_void_p, POINTER(typ))[0] = value

    # see also c:/sf/pywin32/com/win32com/src/oleargs.cpp 54
    def _set_value(self, value):
        _VariantClear(self)
        if value is None:
            self.vt = VT_NULL
        elif (
            hasattr(value, "__len__") and len(value) == 0 and not isinstance(value, str)
        ):
            self.vt = VT_NULL
        # since bool is a subclass of int, this check must come before
        # the check for int
        elif isinstance(value, bool):
            self.vt = VT_BOOL
            self._.VT_BOOL = value
        elif isinstance(value, (int, c_int)):
            self.vt = VT_I4
            self._.VT_I4 = value
        elif isinstance(value, int):
            u = self._
            # try VT_I4 first.
            u.VT_I4 = value
            if u.VT_I4 == value:
                # it did work.
                self.vt = VT_I4
                return
            # try VT_UI4 next.
            if value >= 0:
                u.VT_UI4 = value
                if u.VT_UI4 == value:
                    # did work.
                    self.vt = VT_UI4
                    return
            # try VT_I8 next.
            if value >= 0:
                u.VT_I8 = value
                if u.VT_I8 == value:
                    # did work.
                    self.vt = VT_I8
                    return
            # try VT_UI8 next.
            if value >= 0:
                u.VT_UI8 = value
                if u.VT_UI8 == value:
                    # did work.
                    self.vt = VT_UI8
                    return
            # VT_R8 is last resort.
            self.vt = VT_R8
            u.VT_R8 = float(value)
        elif isinstance(value, (float, c_double)):
            self.vt = VT_R8
            self._.VT_R8 = value
        elif isinstance(value, str):
            self.vt = VT_BSTR
            # do the c_wchar_p auto unicode conversion
            self._.c_void_p = _SysAllocStringLen(value, len(value))
        elif isinstance(value, datetime.datetime):
            delta = value - _com_null_date
            # a day has 24 * 60 * 60 = 86400 seconds
            com_days = (
                delta.days + (delta.seconds + delta.microseconds * 1e-6) / 86400.0
            )
            self.vt = VT_DATE
            self._.VT_R8 = com_days
        elif comtypes.npsupport.isdatetime64(value):
            com_days = value - comtypes.npsupport.com_null_date64
            com_days /= comtypes.npsupport.numpy.timedelta64(1, "D")
            self.vt = VT_DATE
            self._.VT_R8 = com_days
        elif decimal is not None and isinstance(value, decimal.Decimal):
            self._.VT_CY = int(round(value * 10000))
            self.vt = VT_CY
        elif isinstance(value, POINTER(IDispatch)):
            CopyComPointer(value, byref(self._))
            self.vt = VT_DISPATCH
        elif isinstance(value, POINTER(IUnknown)):
            CopyComPointer(value, byref(self._))
            self.vt = VT_UNKNOWN
        elif isinstance(value, (list, tuple)):
            obj = _midlSAFEARRAY(VARIANT).create(value)
            memmove(byref(self._), byref(obj), sizeof(obj))
            self.vt = VT_ARRAY | obj._vartype_
        elif isinstance(value, array.array):
            vartype = _arraycode_to_vartype[value.typecode]
            typ = _vartype_to_ctype[vartype]
            obj = _midlSAFEARRAY(typ).create(value)
            memmove(byref(self._), byref(obj), sizeof(obj))
            self.vt = VT_ARRAY | obj._vartype_
        elif comtypes.npsupport.isndarray(value):
            # Try to convert a simple array of basic types.
            descr = value.dtype.descr[0][1]
            typ = comtypes.npsupport.typecodes.get(descr)
            if typ is None:
                # Try for variant
                obj = _midlSAFEARRAY(VARIANT).create(value)
            else:
                obj = _midlSAFEARRAY(typ).create(value)
            memmove(byref(self._), byref(obj), sizeof(obj))
            self.vt = VT_ARRAY | obj._vartype_
        elif isinstance(value, Structure) and hasattr(value, "_recordinfo_"):
            guids = value._recordinfo_
            from comtypes.typeinfo import GetRecordInfoFromGuids

            ri = GetRecordInfoFromGuids(*guids)
            self.vt = VT_RECORD
            # Assigning a COM pointer to a structure field does NOT
            # call AddRef(), have to call it manually:
            ri.AddRef()
            self._.pRecInfo = ri
            self._.pvRecord = ri.RecordCreateCopy(byref(value))
        elif isinstance(getattr(value, "_comobj", None), POINTER(IDispatch)):
            CopyComPointer(value._comobj, byref(self._))
            self.vt = VT_DISPATCH
        elif isinstance(value, VARIANT):
            _VariantCopy(self, value)
        elif isinstance(value, c_ubyte):
            self._.VT_UI1 = value
            self.vt = VT_UI1
        elif isinstance(value, c_char):
            self._.VT_UI1 = ord(value.value)
            self.vt = VT_UI1
        elif isinstance(value, c_byte):
            self._.VT_I1 = value
            self.vt = VT_I1
        elif isinstance(value, c_ushort):
            self._.VT_UI2 = value
            self.vt = VT_UI2
        elif isinstance(value, c_short):
            self._.VT_I2 = value
            self.vt = VT_I2
        elif isinstance(value, c_uint):
            self.vt = VT_UI4
            self._.VT_UI4 = value
        elif isinstance(value, c_float):
            self.vt = VT_R4
            self._.VT_R4 = value
        elif isinstance(value, c_int64):
            self.vt = VT_I8
            self._.VT_I8 = value
        elif isinstance(value, c_uint64):
            self.vt = VT_UI8
            self._.VT_UI8 = value
        elif isinstance(value, _byref_type):
            ref = value._obj
            self._.c_void_p = addressof(ref)
            self.__keepref = value
            if isinstance(ref, Structure) and hasattr(ref, "_recordinfo_"):
                guids = ref._recordinfo_
                from comtypes.typeinfo import GetRecordInfoFromGuids

                ri = GetRecordInfoFromGuids(*guids)
                self.vt = VT_RECORD | VT_BYREF
                # Assigning a COM pointer to a structure field does NOT
                # call AddRef(), have to call it manually:
                ri.AddRef()
                self._.pRecInfo = ri
                self._.pvRecord = cast(value, c_void_p)
            elif isinstance(ref, _Pointer) and isinstance(
                ref.contents, _safearray.tagSAFEARRAY
            ):
                self.vt = VT_ARRAY | ref._vartype_ | VT_BYREF
                self._.pparray = cast(value, POINTER(POINTER(_safearray.tagSAFEARRAY)))
            else:
                self.vt = _ctype_to_vartype[type(ref)] | VT_BYREF
        elif isinstance(value, _Pointer):
            ref = value.contents
            self._.c_void_p = addressof(ref)
            self.__keepref = value
            if isinstance(ref, Structure) and hasattr(ref, "_recordinfo_"):
                guids = ref._recordinfo_
                from comtypes.typeinfo import GetRecordInfoFromGuids

                ri = GetRecordInfoFromGuids(*guids)
                self.vt = VT_RECORD | VT_BYREF
                # Assigning a COM pointer to a structure field does NOT
                # call AddRef(), have to call it manually:
                ri.AddRef()
                self._.pRecInfo = ri
                self._.pvRecord = cast(value, c_void_p)
            elif isinstance(ref, _safearray.tagSAFEARRAY):
                obj = _midlSAFEARRAY(value._itemtype_).create(value.unpack())
                memmove(byref(self._), byref(obj), sizeof(obj))
                self.vt = VT_ARRAY | obj._vartype_
            elif isinstance(ref, _Pointer) and isinstance(
                ref.contents, _safearray.tagSAFEARRAY
            ):
                self.vt = VT_ARRAY | ref._vartype_ | VT_BYREF
                self._.pparray = cast(value, POINTER(POINTER(_safearray.tagSAFEARRAY)))
            else:
                self.vt = _ctype_to_vartype[type(ref)] | VT_BYREF
        else:
            raise TypeError("Cannot put %r in VARIANT" % value)
        # buffer ->  SAFEARRAY of VT_UI1 ?

    # c:/sf/pywin32/com/win32com/src/oleargs.cpp 197
    def _get_value(self, dynamic=False):
        vt = self.vt
        if vt in (VT_EMPTY, VT_NULL):
            return None
        elif vt == VT_I1:
            return self._.VT_I1
        elif vt == VT_I2:
            return self._.VT_I2
        elif vt == VT_I4:
            return self._.VT_I4
        elif vt == VT_I8:
            return self._.VT_I8
        elif vt == VT_UI8:
            return self._.VT_UI8
        elif vt == VT_INT:
            return self._.VT_INT
        elif vt == VT_UI1:
            return self._.VT_UI1
        elif vt == VT_UI2:
            return self._.VT_UI2
        elif vt == VT_UI4:
            return self._.VT_UI4
        elif vt == VT_UINT:
            return self._.VT_UINT
        elif vt == VT_R4:
            return self._.VT_R4
        elif vt == VT_R8:
            return self._.VT_R8
        elif vt == VT_BOOL:
            return self._.VT_BOOL
        elif vt == VT_BSTR:
            return self._.bstrVal
        elif vt == VT_DATE:
            days = self._.VT_R8
            return datetime.timedelta(days=days) + _com_null_date
        elif vt == VT_CY:
            return self._.VT_CY / decimal.Decimal("10000")
        elif vt == VT_UNKNOWN:
            val = self._.c_void_p
            if not val:
                # We should/could return a NULL COM pointer.
                # But the code generation must be able to construct one
                # from the __repr__ of it.
                return None  # XXX?
            ptr = cast(val, POINTER(IUnknown))
            # cast doesn't call AddRef (it should, imo!)
            ptr.AddRef()
            return ptr.__ctypes_from_outparam__()
        elif vt == VT_DECIMAL:
            return self.decVal.as_decimal()
        elif vt == VT_DISPATCH:
            val = self._.c_void_p
            if not val:
                # See above.
                return None  # XXX?
            ptr = cast(val, POINTER(IDispatch))
            # cast doesn't call AddRef (it should, imo!)
            ptr.AddRef()
            if not dynamic:
                return ptr.__ctypes_from_outparam__()
            else:
                from comtypes.client.dynamic import Dispatch

                return Dispatch(ptr)
        # see also c:/sf/pywin32/com/win32com/src/oleargs.cpp
        elif self.vt & VT_BYREF:
            return self
        elif vt == VT_RECORD:
            from comtypes.client import GetModule
            from comtypes.typeinfo import IRecordInfo

            # Retrieving a COM pointer from a structure field does NOT
            # call AddRef(), have to call it manually:
            punk = self._.pRecInfo
            punk.AddRef()
            ri = punk.QueryInterface(IRecordInfo)

            # find typelib
            tlib = ri.GetTypeInfo().GetContainingTypeLib()[0]

            # load typelib wrapper module
            mod = GetModule(tlib)
            # retrive the type and create an instance
            value = getattr(mod, ri.GetName())()
            # copy data into the instance
            ri.RecordCopy(self._.pvRecord, byref(value))

            return value
        elif self.vt & VT_ARRAY:
            typ = _vartype_to_ctype[self.vt & ~VT_ARRAY]
            return cast(self._.pparray, _midlSAFEARRAY(typ)).unpack()
        else:
            raise NotImplementedError("typecode %d = 0x%x)" % (vt, vt))

    def __getitem__(self, index):
        if index != 0:
            raise IndexError(index)
        if self.vt == VT_BYREF | VT_VARIANT:
            v = VARIANT()
            # apparently VariantCopyInd doesn't work always with
            # VT_BYREF|VT_VARIANT, so do it manually.
            v = cast(self._.c_void_p, POINTER(VARIANT))[0]
            return v.value
        else:
            v = VARIANT()
            _VariantCopyInd(v, self)
            return v.value

    # these are missing:
    # getter[VT_ERROR]
    # getter[VT_ARRAY]
    # getter[VT_BYREF|VT_UI1]
    # getter[VT_BYREF|VT_I2]
    # getter[VT_BYREF|VT_I4]
    # getter[VT_BYREF|VT_R4]
    # getter[VT_BYREF|VT_R8]
    # getter[VT_BYREF|VT_BOOL]
    # getter[VT_BYREF|VT_ERROR]
    # getter[VT_BYREF|VT_CY]
    # getter[VT_BYREF|VT_DATE]
    # getter[VT_BYREF|VT_BSTR]
    # getter[VT_BYREF|VT_UNKNOWN]
    # getter[VT_BYREF|VT_DISPATCH]
    # getter[VT_BYREF|VT_ARRAY]
    # getter[VT_BYREF|VT_VARIANT]
    # getter[VT_BYREF]
    # getter[VT_BYREF|VT_DECIMAL]
    # getter[VT_BYREF|VT_I1]
    # getter[VT_BYREF|VT_UI2]
    # getter[VT_BYREF|VT_UI4]
    # getter[VT_BYREF|VT_INT]
    # getter[VT_BYREF|VT_UINT]

    value = property(_get_value, _set_value)

    def __ctypes_from_outparam__(self):
        # XXX Manual resource management, because of the VARIANT bug:
        result = self.value
        self.value = None
        return result

    def ChangeType(self, typecode):
        _VariantChangeType(self, self, 0, typecode)


VARIANT = tagVARIANT
VARIANTARG = VARIANT

_oleaut32 = OleDLL("oleaut32")

_VariantChangeType = _oleaut32.VariantChangeType
_VariantChangeType.argtypes = (POINTER(VARIANT), POINTER(VARIANT), c_ushort, VARTYPE)

_VariantClear = _oleaut32.VariantClear
_VariantClear.argtypes = (POINTER(VARIANT),)

_SysAllocStringLen = windll.oleaut32.SysAllocStringLen
_SysAllocStringLen.argtypes = c_wchar_p, c_uint
_SysAllocStringLen.restype = c_void_p

_VariantCopy = _oleaut32.VariantCopy
_VariantCopy.argtypes = POINTER(VARIANT), POINTER(VARIANT)

_VariantCopyInd = _oleaut32.VariantCopyInd
_VariantCopyInd.argtypes = POINTER(VARIANT), POINTER(VARIANT)

# some commonly used VARIANT instances
VARIANT.null = VARIANT(None)
VARIANT.empty = VARIANT()
VARIANT.missing = v = VARIANT()
v.vt = VT_ERROR
v._.VT_I4 = 0x80020004
del v

_carg_obj = type(byref(c_int()))
from ctypes import Array as _CArrayType


@comtypes.patcher.Patch(POINTER(VARIANT))
class _(object):
    # Override the default .from_param classmethod of POINTER(VARIANT).
    # This allows to pass values which can be stored in VARIANTs as
    # function parameters declared as POINTER(VARIANT).  See
    # InternetExplorer's Navigate2() method, or Word's Close() method, for
    # examples.
    @classmethod
    def from_param(cls, arg):
        # accept POINTER(VARIANT) instance
        if isinstance(arg, POINTER(VARIANT)):
            return arg
        # accept byref(VARIANT) instance
        if isinstance(arg, _carg_obj) and isinstance(arg._obj, VARIANT):
            return arg
        # accept VARIANT instance
        if isinstance(arg, VARIANT):
            return byref(arg)
        if isinstance(arg, _CArrayType) and arg._type_ is VARIANT:
            # accept array of VARIANTs
            return arg
        # anything else which can be converted to a VARIANT.
        return byref(VARIANT(arg))

    def __setitem__(self, index, value):
        # This is to support the same sematics as a pointer instance:
        # variant[0] = value
        self[index].value = value  # type: ignore


################################################################
# interfaces, structures, ...
class IEnumVARIANT(IUnknown):
    _iid_ = GUID("{00020404-0000-0000-C000-000000000046}")
    _idlflags_ = ["hidden"]
    _dynamic = False

    def __iter__(self):
        return self

    def __next__(self):
        item, fetched = self.Next(1)
        if fetched:
            return item
        raise StopIteration

    def __getitem__(self, index):
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

    def Next(self, celt):
        fetched = c_ulong()
        if celt == 1:
            v = VARIANT()
            self.__com_Next(celt, v, fetched)
            return v._get_value(dynamic=self._dynamic), fetched.value
        array = (VARIANT * celt)()
        self.__com_Next(celt, array, fetched)
        result = [v._get_value(dynamic=self._dynamic) for v in array[: fetched.value]]
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


##from _ctypes import VARIANT_set
##import new
##VARIANT.value = property(VARIANT._get_value, new.instancemethod(VARIANT_set, None, VARIANT))


class tagEXCEPINFO(Structure):
    if TYPE_CHECKING:
        wCode: int
        wReserved: int
        bstrSource: str
        bstrDescription: str
        bstrHelpFile: str
        dwHelpContext: int
        pvReserved: Optional[int]
        pfnDeferredFillIn: Optional[int]
        scode: int

    def __repr__(self):
        return "<EXCEPINFO %s>" % (
            (
                self.wCode,
                self.bstrSource,
                self.bstrDescription,
                self.bstrHelpFile,
                self.dwHelpContext,
                self.pfnDeferredFillIn,
                self.scode,
            ),
        )


tagEXCEPINFO._fields_ = [
    ("wCode", WORD),
    ("wReserved", WORD),
    ("bstrSource", BSTR),
    ("bstrDescription", BSTR),
    ("bstrHelpFile", BSTR),
    ("dwHelpContext", DWORD),
    ("pvReserved", c_void_p),
    # ('pfnDeferredFillIn', WINFUNCTYPE(HRESULT, POINTER(tagEXCEPINFO))),
    ("pfnDeferredFillIn", c_void_p),
    ("scode", SCODE),
]
EXCEPINFO = tagEXCEPINFO


class tagDISPPARAMS(Structure):
    if TYPE_CHECKING:
        rgvarg: Array[VARIANT]
        rgdispidNamedArgs: _Pointer[DISPID]
        cArgs: int
        cNamedArgs: int
    _fields_ = [
        # C:/Programme/gccxml/bin/Vc71/PlatformSDK/oaidl.h 696
        ("rgvarg", POINTER(VARIANTARG)),
        ("rgdispidNamedArgs", POINTER(DISPID)),
        ("cArgs", UINT),
        ("cNamedArgs", UINT),
    ]

    def __del__(self):
        if self._b_needsfree_:
            for i in range(self.cArgs):
                self.rgvarg[i].value = None


DISPPARAMS = tagDISPPARAMS

DISPID_VALUE = 0
DISPID_UNKNOWN = -1
DISPID_PROPERTYPUT = -3
DISPID_NEWENUM = -4
DISPID_EVALUATE = -5
DISPID_CONSTRUCTOR = -6
DISPID_DESTRUCTOR = -7
DISPID_COLLECT = -8


class IDispatch(IUnknown):
    _disp_methods_: ClassVar[List[_DispMemberSpec]]

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
            # but we cannot import comtypes.typeinfo at the top level (recursive imports!).
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

    def GetTypeInfo(self, index: int, lcid: int = 0) -> "hints.ITypeInfo":
        """Return type information.  Index 0 specifies typeinfo for IDispatch"""
        import comtypes.typeinfo

        result = self._GetTypeInfo(index, lcid)  # type: ignore
        return result.QueryInterface(comtypes.typeinfo.ITypeInfo)

    def GetIDsOfNames(self, *names: str, **kw: Any) -> List[int]:
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

    @overload
    def Invoke(
        self, dispid: int, *args: Any, _invkind: int = ..., _lcid: int = ...
    ) -> Any:
        ...  # noqa

    @overload
    def Invoke(
        self,
        dispid: int,
        *args: Any,
        _argspec: Sequence["hints._ArgSpecElmType"],
        _invkind: int = ...,
        _lcid: int = ...,
        **kw: Any,
    ) -> Any:
        ...  # noqa

    def Invoke(self, dispid: int, *args: Any, **kw: Any) -> Any:
        """Invoke a method or property."""

        # Memory management in Dispatch::Invoke calls:
        # https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/passing-parameters
        # Quote:
        #     The *CALLING* code is responsible for releasing all strings and
        #     objects referred to by rgvarg[ ] or placed in *pVarResult.
        #
        # For comtypes this is handled in DISPPARAMS.__del__ and VARIANT.__del__.
        _invkind = kw.pop("_invkind", 1)  # DISPATCH_METHOD
        _lcid = kw.pop("_lcid", 0)
        if kw:
            raise ValueError("named parameters not yet implemented")
        dp = DispParamsGenerator(_invkind).generate(*args)
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
            (hresult, text, details) = err.args
            if hresult == DISP_E_EXCEPTION:
                details = (
                    excepinfo.bstrDescription,
                    excepinfo.bstrSource,
                    excepinfo.bstrHelpFile,
                    excepinfo.dwHelpContext,
                    excepinfo.scode,
                )
                raise COMError(hresult, text, details)
            elif hresult == DISP_E_PARAMNOTFOUND:
                # MSDN says: You get the error DISP_E_PARAMNOTFOUND
                # when you try to set a property and you have not
                # initialized the cNamedArgs and rgdispidNamedArgs
                # elements of your DISPPARAMS structure.
                #
                # So, this looks like a bug.
                raise COMError(hresult, text, argerr.value)
            elif hresult == DISP_E_TYPEMISMATCH:
                # MSDN: One or more of the arguments could not be
                # coerced.
                #
                # Hm, should we raise TypeError, or COMError?
                raise COMError(
                    hresult,
                    text,
                    ("TypeError: Parameter %s" % (argerr.value + 1), args),
                )
            raise
        return result._get_value(dynamic=True)

    # XXX Would separate methods for _METHOD, _PROPERTYGET and _PROPERTYPUT be better?


class DispParamsGenerator(object):
    __slots__ = ("invkind",)

    def __init__(self, invkind: int) -> None:
        self.invkind = invkind

    def generate(self, *args: Any) -> DISPPARAMS:
        """Generate `DISPPARAMS` for passing to `IDispatch::Invoke`.

        Examples:
            >>> _get_rgvarg = lambda dp: [dp.rgvarg[i] for i in range(dp.cArgs)]

            >>> dp = DispParamsGenerator(DISPATCH_METHOD).generate(9)
            >>> _get_rgvarg(dp), bool(dp.rgdispidNamedArgs), dp.cArgs, dp.cNamedArgs
            ([VARIANT(vt=0x3, 9)], False, 1, 0)
            >>> dp = DispParamsGenerator(DISPATCH_PROPERTYGET).generate('foo', 3.14)
            >>> _get_rgvarg(dp), bool(dp.rgdispidNamedArgs), dp.cArgs, dp.cNamedArgs
            ([VARIANT(vt=0x5, 3.14), VARIANT(vt=0x8, 'foo')], False, 2, 0)
            >>> dp = DispParamsGenerator(DISPATCH_PROPERTYPUT).generate(8)
            >>> _get_rgvarg(dp), dp.rgdispidNamedArgs.contents, dp.cArgs, dp.cNamedArgs
            ([VARIANT(vt=0x3, 8)], c_long(-3), 1, 1)
            >>> dp = DispParamsGenerator(DISPATCH_PROPERTYPUTREF).generate(7, 'bar')
            >>> _get_rgvarg(dp), dp.rgdispidNamedArgs.contents, dp.cArgs, dp.cNamedArgs
            ([VARIANT(vt=0x8, 'bar'), VARIANT(vt=0x3, 7)], c_long(-3), 2, 1)

            >>> gen = DispParamsGenerator(DISPATCH_METHOD)
            >>> _get_rgvarg(gen.generate())
            []
            >>> _get_rgvarg(gen.generate(4))
            [VARIANT(vt=0x3, 4)]
            >>> _get_rgvarg(gen.generate(4, 3.14))
            [VARIANT(vt=0x5, 3.14), VARIANT(vt=0x3, 4)]
            >>> _get_rgvarg(gen.generate(4, 3.14, 'foo'))
            [VARIANT(vt=0x8, 'foo'), VARIANT(vt=0x5, 3.14), VARIANT(vt=0x3, 4)]
        """
        array = (VARIANT * len(args))()
        for i, a in enumerate(args[::-1]):
            array[i].value = a
        dp = DISPPARAMS()
        dp.cArgs = len(args)
        if self.invkind in (DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF):  # propput
            dp.cNamedArgs = 1
            dp.rgvarg = array
            dp.rgdispidNamedArgs = pointer(DISPID(DISPID_PROPERTYPUT))
        else:
            dp.cNamedArgs = 0
            dp.rgvarg = array
        return dp


################################################################
# safearrays
# XXX Only one-dimensional arrays are currently implemented

# map ctypes types to VARTYPE values

_arraycode_to_vartype = {
    "d": VT_R8,
    "f": VT_R4,
    "l": VT_I4,
    "i": VT_INT,
    "h": VT_I2,
    "b": VT_I1,
    "I": VT_UINT,
    "L": VT_UI4,
    "H": VT_UI2,
    "B": VT_UI1,
}

_ctype_to_vartype: Dict[Type[_CData], int] = {
    c_byte: VT_I1,
    c_ubyte: VT_UI1,
    c_short: VT_I2,
    c_ushort: VT_UI2,
    c_long: VT_I4,
    c_ulong: VT_UI4,
    c_float: VT_R4,
    c_double: VT_R8,
    c_longlong: VT_I8,
    c_ulonglong: VT_UI8,
    VARIANT_BOOL: VT_BOOL,
    BSTR: VT_BSTR,
    VARIANT: VT_VARIANT,
    # SAFEARRAY(VARIANT *)
    #
    # It is unlear to me if this is allowed or not.  Apparently there
    # are typelibs that define such an argument type, but it may be
    # that these are buggy.
    #
    # Point is that SafeArrayCreateEx(VT_VARIANT|VT_BYREF, ..) fails.
    # The MSDN docs for SafeArrayCreate() have a notice that neither
    # VT_ARRAY not VT_BYREF may be set, this notice is missing however
    # for SafeArrayCreateEx().
    #
    # We have this code here to make sure that comtypes can import
    # such a typelib, although calling ths method will fail because
    # such an array cannot be created.
    POINTER(VARIANT): VT_BYREF | VT_VARIANT,
    # This is needed to import Esri ArcObjects (esriSystem.olb).
    POINTER(BSTR): VT_BYREF | VT_BSTR,
    # These are not yet implemented:
    # POINTER(IUnknown): VT_UNKNOWN,
    # POINTER(IDispatch): VT_DISPATCH,
}

_vartype_to_ctype: Dict[int, Type[_CData]] = {}
for c, v in _ctype_to_vartype.items():
    _vartype_to_ctype[v] = c
_vartype_to_ctype[VT_INT] = _vartype_to_ctype[VT_I4]
_vartype_to_ctype[VT_UINT] = _vartype_to_ctype[VT_UI4]
_ctype_to_vartype[c_char] = VT_UI1


try:
    from comtypes.safearray import _midlSAFEARRAY
except (ImportError, AttributeError):
    pass


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
