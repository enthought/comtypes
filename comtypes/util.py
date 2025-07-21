"""This module defines the funtions byref_at(cobj, offset)
and cast_field(struct, fieldname, fieldtype).
"""

from ctypes import (
    POINTER,
    Structure,
    Union,
    _SimpleCData,
    addressof,
    byref,
    c_byte,
    c_char,
    c_double,
    c_float,
    c_int,
    c_long,
    c_longlong,
    c_short,
    c_void_p,
    cast,
    sizeof,
)
from ctypes import Array as _CArrayType
from typing import TYPE_CHECKING, Type, TypeVar, overload

if TYPE_CHECKING:
    from ctypes import _CArgObject, _CData

_T = TypeVar("_T")
_CT = TypeVar("_CT", bound="_CData")


def _calc_offset():
    # Internal helper function that calculates where the object
    # returned by a byref() call stores the pointer.

    # The definition of PyCArgObject in C code (that is the type of
    # object that a byref() call returns):
    class PyCArgObject(Structure):
        class value(Union):
            _fields_ = [
                ("c", c_char),
                ("h", c_short),
                ("i", c_int),
                ("l", c_long),
                ("q", c_longlong),
                ("d", c_double),
                ("f", c_float),
                ("p", c_void_p),
            ]

        #
        # Thanks to Lenard Lindstrom for this tip:
        # sizeof(PyObject_HEAD) is the same as object.__basicsize__.
        #
        _fields_ = [
            ("PyObject_HEAD", c_byte * object.__basicsize__),
            ("pffi_type", c_void_p),
            ("tag", c_char),
            ("value", value),
            ("obj", c_void_p),
            ("size", c_int),
        ]

        _anonymous_ = ["value"]

    # additional checks to make sure that everything works as expected

    if sizeof(PyCArgObject) != type(byref(c_int())).__basicsize__:
        raise RuntimeError("sizeof(PyCArgObject) invalid")

    obj = c_int()
    ref = byref(obj)

    argobj = PyCArgObject.from_address(id(ref))

    if argobj.obj != id(obj) or argobj.p != addressof(obj) or argobj.tag != b"P":
        raise RuntimeError("PyCArgObject field definitions incorrect")

    return PyCArgObject.p.offset  # offset of the pointer field


################################################################
#
# byref_at
#
@overload
def byref_at(obj: _CArrayType, offset: int) -> "_CArgObject": ...
@overload
def byref_at(obj: "_CData", offset: int) -> "_CArgObject": ...
def byref_at(
    obj,
    offset,
    _byref=byref,
    _c_void_p_from_address=c_void_p.from_address,
    _byref_pointer_offset=_calc_offset(),
):
    """byref_at(cobj, offset) behaves similar this C code:

        (((char *)&obj) + offset)

    In other words, the returned 'pointer' points to the address of
    'cobj' + 'offset'.  'offset' is in units of bytes.
    """
    ref = _byref(obj)
    # Change the pointer field in the created byref object by adding
    # 'offset' to it:
    _c_void_p_from_address(id(ref) + _byref_pointer_offset).value += offset
    return ref


################################################################
#
# cast_field
#
@overload
def cast_field(
    struct: Structure, fieldname: str, fieldtype: Type["_SimpleCData[_T]"]
) -> _T: ...
@overload
def cast_field(struct: Structure, fieldname: str, fieldtype: Type[_CT]) -> _CT: ...
def cast_field(
    struct,
    fieldname,
    fieldtype,
    _POINTER=POINTER,
    _byref_at=byref_at,
):
    """cast_field(struct, fieldname, fieldtype)

    Return the contents of a struct field as it it were of type
    'fieldtype'.
    """
    fieldoffset = getattr(type(struct), fieldname).offset
    return cast(_byref_at(struct, fieldoffset), _POINTER(fieldtype))[0]


__all__ = ["byref_at", "cast_field"]
