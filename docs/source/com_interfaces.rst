###########################
``comtypes`` COM interfaces
###########################

.. contents::

Overview
********

To use or implement a COM interface in |comtypes|, a Python class
must be created. Normally, it is not necessary to write this class
manually since ``comtypes.client.GetModule`` creates interfaces from
type libraries automatically.  However, there may be cases where the
interface is not defined in the available type library, so an
understanding of creating the interface manually and the generated
code is certainly useful.

If no type library but only an IDL file is available it is often the
fastest way to make the interfaces available to Python by compiling
the IDL file into a temporary type library, and generate a Python
module for it; the type library can be deleted after that because it
is not needed any more.

It is possible to take the generated module, move it as a template
into another location and customize the classes with hand written
methods (this is how much of the interfaces in the |comtypes| package
have been created).

The COM interfaces in |comtypes| are abstract classes, they should
never be instantiated.


Defining COM interfaces
***********************

A COM interface in |comtypes| is defined by creating a class.  The
class must derive from ``comtypes.IUnknown`` or a subclass of
``IUnknown``.


The ``IUnknown`` as a Python class
++++++++++++++++++++++++++++++++++

.. py:class:: comtypes.IUnknown

    In this package, ``IUnknown`` is defined as a pure Python class,
    with all its high-level wrapper methods.

    .. py:method:: QueryInterface(interface, iid=None)

        This high-level method wraps the low-level `IUnknown::QueryInterface <https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-queryinterface(refiid_void)>`_
        foreign function calls.  This enables a Pythonic way of usage,
        without worrying about pointers or passing by reference.

        *interface* is a ``IUnknown`` Python class or one of
        subclasses.  If the COM object implements the interface,
        then it returns a pointer instance to that interface after
        incrementing the reference count on it.

        *iid* is the optional interface identifier (IID).  In most
        cases, the *interface* class attribute is used to identify
        the interface.  However, passing a ``comtypes.GUID`` instance
        can be useful in certain low-level processing scenarios.

        The return value is not a ``HRESULT`` value but a pointer to
        the interface.  If the COM object does **not** implement the
        interface, a ``ctypes.COMError`` is raised with an ``hresult``
        of ``-2147467262`` (``E_NOINTERFACE``, ``'0x80004002'`` in
        signed-32bit hex)

    .. py:method:: Add()

        This wraps the `IUnknown::AddRef <https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref>`_.
        Increments the reference count for an interface pointer to a
        COM object and returns the new reference count.

    .. py:method:: Release()

        This wraps the `IUnknown::Release <https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-release>`_.
        Decrements the reference count for an interface on a COM
        object and returns the new reference count.

        In other COM technologies, it is necessary to explicitly
        release COM pointers that have been created or copied by
        calling ``Release``. However, in |comtypes|, explicit release
        is not required because ``Release`` is automatically invoked
        via ``atexit`` hooks or metaclasses when the Python
        interpreter exits or when the Python instance is about to be
        destroyed.

        In fact, explicitly releasing the pointer can cause issues;
        if ``Release`` is called at the aforementioned timing, it may
        raise an ``OSError``.

        .. doctest::

            >>> from comtypes.client import CreateObject, GetModule
            >>> GetModule('UIAutomationCore.dll')  # doctest: +ELLIPSIS
            <module 'comtypes.gen.UIAutomationClient' from ...>
            >>> from comtypes.gen.UIAutomationClient import CUIAutomation
            >>> iuia = CreateObject(CUIAutomation)
            >>> iuia  # doctest: +ELLIPSIS
            <POINTER(IUIAutomation) ptr=... at ...>
            >>> iuia.Release()
            0
            >>> del iuia  # doctest: +ELLIPSIS
            Exception ignored in: <function _compointer_base.__del__ at ...>
            Traceback (most recent call last):
              ...
            OSError: exception: access violation writing ...

    The interface class must define the following class attributes:

    .. py:attribute:: _iid_

        a ``comtypes.GUID`` instance containing the
        *interface identifier* of the interface

    .. py:attribute:: _idlflags_

        (optional) a sequence containing IDL flags for the interface

    .. py:attribute:: _case_insensitive_

        (optional) If set to ``True``, this interface supports case
        insensitive attribute access.

    .. py:attribute:: _methods_

        a sequence describing the methods of this interface.  COM
        methods of the superclass must not be listed, they are
        inherited automatically.

    If one or more of the COM methods reference the interface class
    itself, it is possible to assign the ``_methods_`` attribute
    *after* the class statement like this:

    .. sourcecode:: python

        class ISomeInterface(IUnknown):
            _iid_ = GUID(...)

        ISomeInterface._methods_ = [...,]


The ``_methods_`` list
----------------------

Methods are described in a way that looks somewhat similar to an IDL
definition of a COM interface.  Methods must be listed in VTable
order.

There are two functions that create a method definition: ``STDMETHOD``
is the simple way, and ``COMMETHOD`` allows to specify more
information.

.. py:function:: comtypes.STDMETHOD(restype, methodname, argtypes=())

    Calling ``STDMETHOD`` allows to specify the type of the COM method
    return value.  Usually *restype* is a ``HRESULT``, but other return
    types are also possible.  *methodname* is the name of the COM
    method.  *argtypes* are the types of arguments that the COM
    method expects.


.. py:function:: comtypes.COMMETHOD(idlflags, restype, methodname, *argspec)

    *idlflags* is a list of IDL flags for the method.  Possible values
    include ``dispid(aNumber)`` and ``helpstring(HelpText)``, as well as
    ``"propget"`` for a property getter method, or ``"proput"`` for a
    property setter method.

    *restype* and *methodname* are the same as above.

    *argspec* is a sequence of tuples, each item describing one
    argument for the COM method, and must contain several items:

        1. a sequence of IDL flags: ``"in"``, ``"out"``, ``"retval"``, ``"lcid"``.

        2. type of the argument.

        3. argument name.

..    4. XXX Are there more???

Since the ``IUnknown`` metaclass automatically creates Python methods
and properties that forward the call to the COM methods, there is
typically no need to write any Python methods for the interface class
(unless you want to override what the metaclass does).


An Example
++++++++++

These are two simple COM interfaces. ``IProvideClassInfo`` only
contains one method ``GetClassInfo`` (in addition to the three methods
inherited from ``IUnknown``).  ``IProvideClassInfo2`` inherits from
``IProvideClassInfo`` and adds a ``GetGUID`` method.

This is the IDL definition, slightly simplified (from Microsoft's
``OCIDL.IDL``):

.. sourcecode:: idl

    [
        object,
        uuid(B196B283-BAB4-101A-B69C-00AA00341D07),
        pointer_default(unique)
    ]
    interface IProvideClassInfo : IUnknown
    {
        HRESULT GetClassInfo(
                    [out] ITypeInfo ** ppTI
                );
    }

    [
        object,
        uuid(A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851),
        pointer_default(unique)
    ]
    interface IProvideClassInfo2 : IProvideClassInfo
    {
        HRESULT GetGUID(
                    [in]  DWORD dwGuidKind,
                    [out] GUID * pGUID
                );
    }


|comtypes| interface classes:

.. sourcecode:: python

    from ctypes import *
    from comtypes import IUnknown, GUID, COMMETHOD
    from comtypes.typeinfo import ITypeInfo

    class IProvideClassInfo(IUnknown):
        _iid_ = GUID("{B196B283-BAB4-101A-B69C-00AA00341D07}")
        _methods_ = [
            COMMETHOD([], HRESULT, "GetClassInfo",
                      ( ['out'],  POINTER(POINTER(ITypeInfo)), "ppTI" ) )
            ]

    class IProvideClassInfo2(IProvideClassInfo):
        _iid_ = GUID("{A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851}")
        _methods_ = [
            COMMETHOD([], HRESULT, "GetGUID",
                      ( ['in'], DWORD, "dwGuidKind" ),
                      ( ['out', 'retval'], POINTER(GUID), "pGUID" ))
            ]


Using COM interfaces
********************

As said above, |comtypes| interface classes are never instantiated,
also they are never used directly.  Instead, one uses instances of
``POINTER(ISomeInterface)`` to call the methods on a COM object.

The ``IUnknown`` COM interface has ``AddRef()``, ``Release()``, and
``QueryInterface()`` methods that you can call.  Since the COM internal
reference count is handled automatically by |comtypes|, there is no
need to call the first two methods.

``QueryInterface()``, however, is the call that you need to ask a COM
object for other COM interfaces.  Since IUnknown is the base class of
**all** COM interfaces, it is available in every COM interface.

So, assuming you have a ``POINTER(IUnknown)`` instance, you can ask
for another interface by calling ``QueryInterface`` with the interface
you want to use.  For example:

.. sourcecode:: python

    # punk is a pointer to an IUnknown interface
    pci = punk.QueryInterface(IProvideClassInfo)


This call will either succeed and return a
``POINTER(IProvideClassInfo)`` instance, or it will raise a
``comtypes.COMError`` if the interface is not supported.  Assuming the
call succeeded, you can get the type information of the object by
calling:

.. sourcecode:: python

    ti = pci.GetClassInfo()


Unless the call fails, it will return a ``POINTER(ITypeInfo)``
instance.


Implementing COM interfaces
***************************

While the ``IUnknown`` metaclass creates Python methods that you can
call in client code directly, you have to write code yourself if you
want to **implement** a COM interface.  One important thing to keep
in mind is that each COM method implementation with |comtypes|
receives an additional special parameter per convention named
*this*, just after the *self* standard parameter.

If you want to implement the ``IProvideClassInfo`` interface described
above in a Python class you have to write an implementation of the
``GetClassInfo`` method:

.. sourcecode:: python

    from comtypes import COMObject
    from comtypes.persist import IProvideClassInfo

    class MyCOMObject(COMObject):
        _com_interfaces_ = [
            ...,
            IProvideClassInfo,
        ]


Skipping some very important details that are out of context here, the
interfaces that your COM object implements must be listed in the
``_com_interfaces_`` class variable.  Then, of course, you should
implement the methods of all the interfaces by writing a Python method
for each of them.

.. note::

    The ``COMObject`` metaclass provides a default for methods
    that are **not** implemented in Python.  This default method returns
    the standard COM error code ``E_NOTIMPL`` when it is called.

To implement the COM method named ``MethodName`` for the interface
``ISomeInterface`` you write a Python method either named ``ISomeInterface_MethodName``
or simply ``MethodName``.

This method must accept the following arguments:

  1. the standard Python ``self`` parameter.

  2. a special *this* parameter, that you can usually ignore.

  3. All the parameters that are listed in the interface description.

The latter parameters will be instances of types specified in the
``_methods_`` description.

So, to implement the ``GetClassInfo`` method of the
``IProvideClassInfo`` interface, one could write this code:

.. sourcecode:: python

    from comtypes import COMObject
    from comtypes.persist import IProvideClassInfo

    class MyCOMObject(COMObject):
        _com_interfaces_ = [
            ...,
            IProvideClassInfo,
        ]

        def IProvideClassInfo_GetClassInfo(self, this, ppTI):
	        ...  # this method could also be named 'GetClassInfo'.


The *ppTI* parameter in this case is an instance of
``POINTER(POINTER(ITypeInfo))`` which you have to fill out.  So, to
write a method that actually returns a useful type info pointer for
the object, you have to fill the contents of the *ppTI* pointer like
this:

.. sourcecode:: python

    def IProvideClassInfo_GetClassInfo(self, this, ppTI):
        from comtypes.hresult import E_POINTER, S_OK
        # First, check for NULL pointer and return error
        if not ppTI:
            return E_POINTER
        ti = create_type_info(...) # get the type info somehow
        # poke it into the 'out' parameter
        ppTI[0] = ti
        # and return success
        return S_OK


``E_POINTER`` is an error code that you should return when you
received an unexpected NULL pointer, ``S_OK`` is the usual success
code for COM methods returning a ``HRESULT``.  For details about the
semantics that you have to implement for a COM interface method
consult the MSDN documentation.


Case sensitivity
****************

In principle, COM is a case insensitive technology (probably because
of Visual Basic).  Type libraries generated from IDL files, however,
do *not* always even preserve the case of identifiers; see for example
http://support.microsoft.com/kb/220137.

Python (and C/C++) are case sensitive languages, so |comtypes| is
also case sensitive.  This means that you have to call
``obj.QueryInterface(...)``, it will not work to write
``obj.queryinterface(...)``.

To work around the problems that you get when the case of identifiers
in the type library (and in the generated Python module for this
library) is not the same as in the IDL file, |comtypes| allows to
have case insensitive attribute access for methods and properties of
COM interfaces.  This behaviour is enabled by setting the
``_case_insensitive_`` attribute of a Python COM interface to
``True``.  In case of derived COM interfaces, case sensitivity is
enabled or disabled separately for each interface.

The code generated by the ``GetModule`` function sets this attribute
to ``True``.  Case insensitive access has a small performance penalty,
if you want to avoid this, you should edit the generated code and set
``_case_insensitive_`` to False.


More about the metaclass
************************

The Python class ``IUnknown``, which is the base interface of *all*
COM interfaces, uses a metaclass that automatically creates Python
methods and properties for the COM methods described in the
``_methods_`` list.

For a COM method described by a ``STDMETHOD`` only the types of the
arguments and the return type of the method is known.  In this case
only trivial code is generated that checks the type of the arguments
and returns whatever the COM method returns.

For a COM method described by ``COMMETHOD``, much more information is
available: the argument names, the direction of data transfer for each
argument ``["in"]``, ``["out"]``, or ``["in", "out"]``, and whether
this method is a getter or setter of a property.  In this case, code
is generated that instantiates containers for "out" parameters inside
the method call, passes and ``"in"`` and ``"out"`` parameters to the
actual COM method of the object, retrives ``"out"`` parameters from
their container(s) and returns them as the result.  If the method has
exactly one ``"out"`` parameter, this is returned. If the method has
two or more ``"out"`` parameters, a tuple of their values is returned.

.. note::

    The native return value of the method, usually a ``HRESULT``,
    is **not** returned in the presence of "out" parameters.

For the ``IProvideClassInfo`` and ``IProvideClassInfo`` COM interfaces
mentioned above, the metaclass creates methods with these signatures
automatically (``__call_com_method()`` is the ``ctypes`` code that
calls the actual method slot of the COM object):

.. sourcecode:: python

    class IProvideClassInfo(IUnknown):
        _iid_ = GUID("{B196B283-BAB4-101A-B69C-00AA00341D07}")
        # code for this method generated by the IUnknown metaclass at
        # runtime
        # def GetClassInfo(self):
        #     param = POINTER(ITypeInfo)()
        #     __call_com_method(byref(param))
        #     return param[0]

    class IProvideClassInfo2(IProvideClassInfo):
        _iid_ = GUID("{A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851}")
        # code for this method generated by the IUnknown metaclass at
        # runtime
        # def GetGUID(self, dwGuidKind):
        #     param = GUID()
        #     __call_com_method(dwGuidKind, byref(param))
        #     return param


According to MSDN, the ``IProvideClassInfo2::GetGUID`` method
*"returns a GUID corresponding to the specified dwGuidKind"*.
However, currently only a single valid value for *dwGuidKind* is
defined: ``GUIDKIND_DEFAULT_SOURCE_DISP_IID == 1`` which specifies the guid
for the default outgoing interface.

So, it would probably make sense to implement the GetGUID method with
a default value of 1 for the *dwGuidKind* parameter.  This can be done
by manually implementing a ``GetGUID`` method for the
``IProvideClassInfo2`` interface class:

.. sourcecode:: python

    class IProvideClassInfo2(IProvideClassInfo):
        ...
        def GetGUID(self, dwGuidKind=1):
            return self._GetGUID(dwGuidKind)


When the metaclass finds that the ``GetGUID`` method **already has**
an implementation, it will not overwrite it.  Instead, it creates an
interface method with the name ``_GetGUID`` that you can use to get
the raw functionality.


.. |comtypes| replace:: ``comtypes``
