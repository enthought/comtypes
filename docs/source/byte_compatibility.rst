Signed 8-bit values from type libraries
=======================================

This page describes the compatibility impact of the proposed ``VT_I1``
type-library mapping change. The change is under review; the examples below
describe regenerated modules using that mapping.
The decision and compatibility discussion is tracked in
`issue #969 <https://github.com/enthought/comtypes/issues/969>`_.

Generated types
---------------

``GetModule`` now represents ``VT_I1`` with ``ctypes.c_byte`` (a signed 8-bit
integer) instead of ``ctypes.c_char``. This affects scalar values, pointers,
structure and union fields, and SAFEARRAY elements described by ``VT_I1``
in a type library. For example,
an IDL ``small`` scalar and a plain ``char*`` pointer use this mapping.
Explicit string types such as ``LPSTR`` still generate ``STRING``
(``ctypes.c_char_p``); ``VT_UI1`` still generates ``ctypes.c_ubyte``.

The change applies when modules are generated again. Existing generated
source files and interfaces defined directly in Python retain their declared
types. Mixing old generated modules with new ones can therefore produce
different Python argument and return types for the same type library.

Arguments
---------

Pass a Python integer in the range -128 to 127, or a ``ctypes.c_byte`` value,
to a signed 8-bit scalar argument. The generated retval for a signed scalar
is also numeric: an ``EchoInt8(-7)`` call returns the integer ``-7``.
Single-character ``bytes`` and ``ctypes.c_char`` arguments do not match the
new scalar type.

A plain ``char*`` argument previously generated as ``STRING`` now generates
as ``POINTER(c_byte)``. Existing calls that pass ``b"h"`` to such an argument
raise ``ctypes.ArgumentError``. Use a pointer, a reference, or an array of
``c_byte`` values instead, according to the native method's buffer contract::

    from ctypes import byref, c_byte, pointer

    value = c_byte(ord("h"))
    server.ReadCharPointer(pointer(value))
    server.ReadCharPointer(byref(value))
    server.ReadCharPointer((c_byte * 1)(ord("h")))

These examples refer to a method that reads exactly one character. For
methods that read more data, provide the required buffer size and lifetime.
The pointer type alone does not describe a string's length or termination.
Explicit ``LPSTR`` arguments continue to accept ``bytes`` as before.

Return values
-------------

A plain ``char*`` output also changes from ``STRING`` to ``POINTER(c_byte)``.
For a method returning one character through that parameter, a generated
high-level call now returns an integer. The value 104 returns as 104, and
255 returns as -1. Under the old mapping the automatic output allocation
could instead fail with ``TypeError: c_char_p 'out' parameter must be passed
as default value``.

Calls to the generated raw COM method previously accepted a writable
``ctypes.create_string_buffer`` for that output. The new pointer type rejects
that character array. Supply ``byref(c_byte())`` or a ``c_byte`` array for a
single-character output, and read the resulting signed number. To recover
the corresponding one-byte ``bytes`` value, use ``bytes([value & 0xff])``.

For larger returned buffers, read the result using the native method's
documented length. ``ctypes.string_at(result, length)`` makes a ``bytes`` copy
from a pointer of known length. Use NUL-terminated string conversion only
when the native contract guarantees termination. A plain character pointer
does not guarantee a terminator. Free or release native allocations as
specified by that contract after copying the data.

Explicit ``LPSTR`` output types retain their existing string conversion.
Changing interface declarations by hand should likewise follow the actual
native contract. Type libraries describe both numeric ``small`` and plain
character data using ``VT_I1``, so this mapping change cannot by itself
preserve every existing implicit string conversion.
