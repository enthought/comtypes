#############
NumPy interop
#############

NumPy provides the *de facto* array standard for Python. Though NumPy
is not required to use |comtypes|, |comtypes| provides various options for
NumPy interoperability. NumPy version 1.7 or greater is required to access
all of these features.


.. contents::

NumPy Arrays as Input Arguments
*******************************

NumPy arrays can be passed as ``VARIANT`` arrays arguments. The array is
converted to a SAFEARRAY according to its type. The type conversion
is defined by the ``numpy.ctypeslib`` module.  The following table
shows type conversions that can be performed quickly by (nearly) direct
conversion of a numpy array to a SAFEARRAY. Arrays with type that do not
appear in this table, including object arrays, can still be converted to
SAFEARRAYs on an item-by-item basis.

+------------------------------------------------+---------------+
| NumPy type                                     | VARIANT type  |
+================================================+===============+
| ``int8``                                       | VT_I1         |
+------------------------------------------------+---------------+
| ``int16``, ``short``                           | VT_I2         |
+------------------------------------------------+---------------+
| ``int32``, ``int``, ``intc``, ``int_``         | VT_I4         |
+------------------------------------------------+---------------+
| ``int64``, ``long``, ``longlong``, ``intp``    | VT_I8         |
+------------------------------------------------+---------------+
| ``uint8``, ``ubyte``                           | VT_UI1        |
+------------------------------------------------+---------------+
| ``uint16``, ``ushort``                         | VT_UI2        |
+------------------------------------------------+---------------+
| ``uint32``, ``uint``, ``uintc``                | VT_UI4        |
+------------------------------------------------+---------------+
| ``uint64``, ``ulonglong``, ``uintp``           | VT_UI8        |
+------------------------------------------------+---------------+
| ``float32``                                    | VT_R4         |
+------------------------------------------------+---------------+
| ``float64``, ``float_``                        | VT_R8         |
+------------------------------------------------+---------------+
| ``datetime64``                                 | VT_DATE       |
+------------------------------------------------+---------------+

NumPy Arrays as Output Arguments
********************************

By default, |comtypes| converts SAFEARRAY output arguments to tuples of
python objects on an item-by-item basis.  When dealing with large
SAFEARRAYs, this conversion can be costly.  Comtypes provides a the
``safearray_as_ndarray`` context manager (from ``comtypes.safearray``)
for modifying this behavior to return a NumPy array. This altered
behavior is to put an ndarray over a copy of the SAFEARRAY's memory,
which is faster than calling into python for each item. When this fails,
a NumPy array can still be created on an item-by-item basis.  The context
manager is thread-safe, in that usage of the context manager on one
thread does not affect behavior on other threads.

This is a hypothetical example of using the context manager. The context
manager can be used around any property or method call to retrieve a
NumPy array rather than a tuple.


.. sourcecode:: python

    """Sample demonstrating use of safearray_as_ndarray context manager """

    from comtypes.safearray import safearray_as_ndarray

    # Hypothetically, this returns a SAFEARRAY as a tuple
    data1 = some_interface.some_property

    # This will return a NumPy array, and will be faster for basic types.
    with safearray_as_ndarray:
        data2 = some_interface.some_property


.. |comtypes| replace:: ``comtypes``
