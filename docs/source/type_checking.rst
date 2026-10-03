################################
Type Checking and Type Inference
################################

``comtypes`` generates Python wrapper module files from COM type
libraries.
These generated modules contain extensive type hints which allow
static type checkers to infer the shapes of COM objects, methods,
properties and the behaviour of factories such as ``CreateObject``
or ``GetModule``.

.. contents::

Basic example – ``Scripting.Dictionary``
****************************************

.. sourcecode:: python

    import sys
    from ctypes import POINTER

    from comtypes.client import CreateObject, GetModule

    if sys.version_info >= (3, 11):
        from typing import assert_type
    else:
        from typing_extensions import assert_type

    # Generate/ensure the existence of the ``Scripting`` module.
    GetModule("scrrun.dll")
    # Must be imported statically;
    # The type checker cannot perform static type analysis on the
    # ``ModuleType`` instance returned by ``GetModule``.
    from comtypes.gen import Scripting


    dic = CreateObject(
        Scripting.Dictionary, interface=Scripting.IDictionary
    )
    # At runtime, ``dic`` is a ``POINTER(IDictionary)``.
    # It behaves as a subclass of ``IDictionary`` via its metaclass.
    # Since capabilities to express dynamic subclasses created by
    # ``POINTER`` are not introduced into Python's static type system
    # yet, it is typed to return ``IDictionary``.
    # The static type checker sees the following type:
    #   dic: Scripting.IDictionary
    assert isinstance(dic, POINTER(Scripting.IDictionary))
    assert isinstance(dic, Scripting.IDictionary)
    assert_type(dic, Scripting.IDictionary)

    # Properties are correctly typed
    dic.CompareMode = Scripting.TextCompare
    # This test verifies that the ``Scripting.Dictionary`` supports
    # both subscriptable and callable access patterns, and that the
    # static type checker correctly recognizes these operations.
    dic["foo"] = 1
    assert dic["foo"] == dic.Item["foo"] == 1
    dic.Add("bar", 2)
    assert dic("bar") == dic.Item("bar") == 2
    dic.Item["qux"] = 3
    assert dic("qux") == dic.Item("qux") == 3


Limitations
***********

The COM factory is annotated as returning an ``IUnknown``-based
object, although the actual runtime type is the dynamically defined
subclass ``POINTER(IUnknown)``, created through the interaction
between the ``ctypes.POINTER`` factory function and the complex
metaclass of ``IUnknown``.

It is not annotated as the commonly used ctypes pointer type
``ctypes._Pointer[IUnknown]`` because Python's type system does not
currently implement the intersection types needed to represent such
a dynamically defined subclass. Consequently,
``ctypes._Pointer[IDictionary]`` is interpreted as a container type,
preventing static analysis tools from recognizing the COM pointer
methods.

The recommended typing style in Python is to annotate return values
with the base or abstract interface that describes how the value is
expected to be used, rather than with its concrete runtime type.
For example, if a value is expected to be used only as a sequence,
without adding elements to it, ``collections.abc.Sequence[str]`` is
preferable to ``list[str]``, even when the actual returned value is
a ``list[str]``.
By annotating the return value as ``IUnknown``, this base-class
approach enables static analysis tools to recognize COM interface
methods that cannot be recognized with a
``ctypes._Pointer[IUnknown]`` annotation.
