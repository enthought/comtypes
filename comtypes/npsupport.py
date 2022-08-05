""" Consolidation of numpy support utilities. """
import sys

is_64bits = sys.maxsize > 2**32
HAVE_NUMPY = False
com_null_date64 = None
numpy = None
datetime64 = None
VARIANT_dtype = None
typecodes = {}

def _make_variant_dtype():
    """ Create a dtype for VARIANT. This requires support for Unions, which is
    available in numpy version 1.7 or greater.

    This does not support the decimal type.

    Returns None if the dtype cannot be created.

    """
    numpy = get_numpy()

    # pointer typecode
    ptr_typecode = '<u8' if is_64bits else '<u4'

    _tagBRECORD_format = [
        ('pvRecord', ptr_typecode),
        ('pRecInfo', ptr_typecode),
    ]

    # overlapping typecodes only allowed in numpy version 1.7 or greater
    U_VARIANT_format = dict(
        names=[
            'VT_BOOL', 'VT_I1', 'VT_I2', 'VT_I4', 'VT_I8', 'VT_INT', 'VT_UI1',
            'VT_UI2', 'VT_UI4', 'VT_UI8', 'VT_UINT', 'VT_R4', 'VT_R8', 'VT_CY',
            'c_wchar_p', 'c_void_p', 'pparray', 'bstrVal', '_tagBRECORD',
        ],
        formats=[
            '<i2', '<i1', '<i2', '<i4', '<i8', '<i4', '<u1', '<u2', '<u4',
            '<u8', '<u4', '<f4', '<f8', '<i8', ptr_typecode, ptr_typecode,
            ptr_typecode, ptr_typecode, _tagBRECORD_format,
        ],
        offsets=[0] * 19  # This is what makes it a union
    )

    tagVARIANT_format = [
        ("vt", '<u2'),
        ("wReserved1", '<u2'),
        ("wReserved2", '<u2'),
        ("wReserved3", '<u2'),
        ("_", U_VARIANT_format),
    ]

    return numpy.dtype(tagVARIANT_format)


def isndarray(value):
    """ Check if a value is an ndarray.

    This cannot succeed if numpy is not available.

    """
    if not HAVE_NUMPY:
        return False
    numpy = get_numpy()
    return isinstance(value, numpy.ndarray)


def isdatetime64(value):
    """ Check if a value is a datetime64.

    This cannot succeed if datetime64 is not available.

    """
    if not HAVE_NUMPY:
        return False
    return isinstance(value, datetime64)


def _check_ctypeslib_typecodes():
    import numpy as np
    from numpy import ctypeslib
    try:
        from numpy.ctypeslib import _typecodes
    except ImportError:
        from numpy.ctypeslib import as_ctypes_type

        dtypes_to_ctypes = {}

        for tp in set(np.sctypeDict.values()):
            try:
                ctype_for = as_ctypes_type(tp)
                dtypes_to_ctypes[np.dtype(tp).str] = ctype_for
            except NotImplementedError:
                continue
        ctypeslib._typecodes = dtypes_to_ctypes
    return ctypeslib._typecodes


def enable_numpy_interop():
    """ Import the numpy library (if not already imported) and set up the
    necessary functions for comtypes to work with ndarrays.

    """
    global numpy
    import numpy

    global HAVE_NUMPY
    HAVE_NUMPY = True
    global typecodes
    typecodes = _check_ctypeslib_typecodes()
    # dtype for VARIANT. This allows for packing of variants into an array, and
    # subsequent conversion to a multi-dimensional safearray.
    global VARIANT_dtype
    try:
        VARIANT_dtype = _make_variant_dtype()
    except ValueError:
        pass
    global datetime64
    global com_null_date64
    # This simplifies dependent modules
    try:
        from numpy import datetime64
    except ImportError:
        pass
    else:
        try:
            # This does not work on numpy 1.6
            com_null_date64 = datetime64("1899-12-30T00:00:00", "ns")
        except TypeError:
            pass


def get_numpy():
    """ Returns the numpy package if numpy interop is enabled, otherwise raises
    an import error to remind the user they need to manually enable numpy
    interop

    """
    if HAVE_NUMPY:
        return numpy
    raise ImportError(
        "In comtypes>=1.2.0 numpy interop must be explicitly enabled with "
        "comtypes.npsupport.enable_numpy_interop before attempting to use "
        "numpy features."
    )
