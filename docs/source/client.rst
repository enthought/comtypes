###############################
The ``comtypes.client`` package
###############################

The ``comtypes.client`` package implements the high-level |comtypes|
functionality.

.. contents::

Creating and accessing COM objects
**********************************

``comtypes.client`` exposes three functions that allow to create or
access COM objects.

``CreateObject(progid, clsctx=None, machine=None, interface=None, dynamic=False, pServerInfo=None)``
    Create a COM object and return an interface pointer to it.

    ``progid`` specifies which object to create.  It can be a string
    like ``"InternetExplorer.Application"`` or
    ``"{2F7860A2-1473-4D75-827D-6C4E27600CAC}"``, a ``comtypes.GUID``
    instance, or any object with a ``_clsid_`` attribute that must be
    a ``comtypes.GUID`` instance or a GUID string.

    ``clsctx`` specifies how to create the object, any combination of
    the ``comtypes.CLSCTX_...`` constants can be used.  If nothing is
    passed, ``comtypes.CLSCTX_SERVER`` is used.

    ``machine`` allows to specify that the object should be created on
    a different machine, it must be a string specifying the computer
    name or IP address.  DCOM must be enabled for this to work.

    ``interface`` specifies the interface class that should be
    returned, if not specified |comtypes| will determine a useful
    interface itself and return a pointer to that.

    ``dynamic`` specifies that the generated interface should use
    dynamic dispatch. This is only available for automation interfaces
    and does not generate typelib wrapper.

    ``pServerInfo`` that allows you to specify more information about
    the remote machine than the ``machine`` parameter. It is a pointer
    to a ``COSERVERINFO``. ``machine`` and ``pServerInfo`` may not be
    simultaneously supplied.  DCOM must be enabled for this to work.

``CoGetObject(displayname, interface=None)``
    Create a named COM object and returns an interface pointer to it.
    For the interpretation of ``displayname`` consult the Microsoft
    documentation for the Windows ``CoGetObject`` function.
    ``"winmgmts:"``, for example, is the displayname for `WMI
    monikers`_:

    .. sourcecode:: python

        wmi = CoGetObject("winmgmts:")

    ``interface`` and ``dynamic`` have the same meaning as in the
    ``CreateObject`` function.

``GetActiveObject(progid, interface=None)``
    Returns a pointer to an already running object.  ``progid``
    specifies the active object from the OLE registration database.

    The ``GetActiveObject`` function succeeds when the COM object is
    already running, and has registered itself in the COM running
    object table.  Not all COM objects do this. The arguments are as
    described under ``CreateObject``.

All the three functions mentioned above will create the typelib
wrapper automatically if the object provides type information.  If the
type library is not exposed by the object itself, the wrapper can be
created by calling the ``GetModule`` function.


Using COM objects
*****************

The COM interface pointer that is returned by one of the creation
functions (``CreateObject``, ``CoGetObject``, or ``GetActiveObject``)
exposes methods and properties of the interface (unless ``dynamic``
is passed to the function).

Since |comtypes| uses early binding to COM interfaces (when type
information is exposed by the COM object), the interface methods and
properties are available for introspection.  The Python builtin
``help`` function can be used to get an overview of them.

``Scripting.FileSystemObject`` is the progid of the Microsoft Scripting
Runtime's FileSystemObject; this COM object provides access to the
computer's file system, allowing scripts to create, read, update, and
delete files and folders.

.. doctest::

    >>> from comtypes.client import CreateObject
    >>> fso = CreateObject("Scripting.FileSystemObject")
    >>> help(fso)  # doctest: +ELLIPSIS
    Help on POINTER(IFileSystem...


Calling methods
+++++++++++++++

Calling COM methods is straightforward just like with other Python
objects.  They can be called with positional and named arguments.

Arguments marked ``[out]`` or ``[out, retval]`` in the IDL are
returned from a sucessful method call, in a tuple if there is more
than one.  If no ``[out]`` or ``[out, retval]`` arguments are present,
the ``HRESULT`` returned by the method call is returned.  When
``[out]`` or ``[out, retval]`` arguments are returned from a sucessful
call, the ``HRESULT`` value is lost.

If the COM method call fails, a ``COMError`` exception is raised,
containing the ``HRESULT`` value.


Accessing properties
++++++++++++++++++++

COM properties present some challenges.  Properties can be read-write,
read-only, or write-only.  They may have zero, one, or more arguments;
arguments may even be optional.

The ``Scripting.Dictionary`` object provides a dictionary-like interface.
This example demonstrates accessing and modifying the ``CompareMode``
property, which controls how keys are compared:

.. doctest::

    >>> dic = CreateObject("Scripting.Dictionary")
    >>> dic.CompareMode  # default is 0, BinaryCompare
    0
    >>> dic.CompareMode = 1  # TextCompare
    >>> dic.CompareMode
    1


Properties with arguments (named properties)
--------------------------------------------

Properties with arguments can be accessed using index notation.
The following example starts Excel, creates a new workbook, and
accesses the contents of some cells in the ``xlRangeValueDefault``
format (this code has been tested with version 2402 build
16.0.17328.20670):

.. doctest::
    :skipif: NO_EXCEL

    >>> xl = CreateObject('Excel.Application')
    >>> xl.Workbooks.Add()  # doctest: +ELLIPSIS
    <POINTER(_Workbook) ptr=... at ...>
    >>> from comtypes.gen.Excel import xlRangeValueDefault
    >>> xl.Range["A1", "C1"].Value[xlRangeValueDefault] = (10,'20',31.4)
    >>> xl.Range["A1", "C1"].Value[xlRangeValueDefault]
    ((10.0, 20.0, 31.4),)


Properties with optional arguments
----------------------------------

If you look into the Excel type library (or the generated
``comtypes.gen`` wrapper module) you will find that the parameter for
the ``.Value`` property is optional, so it would be possible to get or
set this property without the need to pass (or even know) the
``xlRangeValueDefault`` argument.

Unfortunately, Python does not allow indexing without arguments:

.. doctest::
    :skipif: NO_EXCEL

    >>> xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
    Traceback (most recent call last):
      ...
        xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
                                   ^
    SyntaxError: invalid syntax
    >>> print(xl.Range["A1", "C1"].Value[])
    Traceback (most recent call last):
      ...
        print(xl.Range["A1", "C1"].Value[])
              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    SyntaxError: invalid syntax. Perhaps you forgot a comma?


So, |comtypes| must provide some ways to access these properties.  To
*get* a named property without passing any argument, you can *call*
the property:

.. doctest::
    :skipif: NO_EXCEL

    >>> print(xl.Range["A1", "C1"].Value())
    ((10.0, 20.0, 31.4),)


It is also possible to index with an empty slice or empty tuple:

.. doctest::
    :skipif: NO_EXCEL

    >>> print(xl.Range["A1", "C1"].Value[:])
    ((10.0, 20.0, 31.4),)
    >>> print(xl.Range["A1", "C1"].Value[()])
    ((10.0, 20.0, 31.4),)


To *set* a named property without passing any argument, you can
also use the empty slice or tuple index trick:

.. doctest::
    :skipif: NO_EXCEL

    >>> xl.Range["A1", "C1"].Value[:] = (3, 2, 1)
    >>> print(xl.Range["A1", "C1"].Value[:])
    ((3.0, 2.0, 1.0),)
    >>> xl.Range["A1", "C1"].Value[()] = (1, 2, 3)
    >>> print(xl.Range["A1", "C1"].Value[()])
    ((1.0, 2.0, 3.0),)


.. This is not (yet?) implemented.  Would is be useful?
   Another way is to assing to the tuple in the normal way:

      >>> xl.Range["A1", "C1"].Value = (1, 2, 3)
      >>>

The lcid parameter
++++++++++++++++++

Some COM methods or properties have an optional ``lcid`` parameter.
This parameter is used to specify a langauge identifier.  The
generated modules always pass 0 (zero) for this parameter.  If this is
not what you want you have to edit the generated code.

Converting data types
+++++++++++++++++++++

|comtypes| usually converts arguments and results between COM and
Python in just the way one would expect.

``VARIANT`` parameters sometimes requires special care.  A ``VARIANT``
can hold a lot of different types - simple ones like integers, floats,
or strings, also more complicated ones like single dimensional or even
multidimensional arrays.  The value a ``VARIANT`` contains is
specified by a *typecode* that |comtypes| automatically assigns.

When you pass simple sequences (lists or tuples) as ``VARIANT``
parameters, the COM server will receive a ``VARIANT`` containing
a ``SAFEARRAY`` of VARIANTs with the typecode ``VT_ARRAY | VT_VARIANT``.

Some COM server methods, however, do not accept such arrays, they
require for example an array of short integers with the typecode
``VT_ARRAY | VT_I2``, an array of integers with typecode ``VT_ARRAY |
VT_INT``, or an array a strings with typecode ``VT_ARRAY | VT_BSTR``.

To create these variants you must pass an instance of the Python
``array.array`` with the correct Python typecode to the COM method.
Note that NumPy arrays are also an option here, as is described in
the following section.

The mapping of the ``array.array`` typecode to the ``VARIANT``
typecode is defined in the ``comtypes.automation`` module by a
dictionary:

.. sourcecode:: python

    _arraycode_to_vartype = {
        "b": VT_I1,
        "h": VT_I2,
        "i": VT_INT,
        "l": VT_I4,

        "B": VT_UI1,
        "H": VT_UI2,
        "I": VT_UINT,
        "L": VT_UI4,

        "f": VT_R4,
        "d": VT_R8,
    }

AutoCAD, for example, is one of the COM servers that requires VARIANTs
with the typecodes ``VT_ARRAY | VT_I2`` or ``VT_ARRAY | VT_R8`` for
parameters.  This code snippet was contributed by a user:

.. sourcecode:: python

    """Sample to demonstrate how to use comtypes to automate AutoCAD:
    adding a point and a line to the drawing; and attaching xdata of
    different types to them. The objective is to actually show how to
    create variants of different types using comtypes.  Such variants are
    required by many methods in AutoCAD COM API. AutoCAD needs to be
    running to test the following code."""
   
    import array
    import comtypes.client
   
    #Get running instance of the AutoCAD application
    app = comtypes.client.GetActiveObject("AutoCAD.Application")
   
    #Get the ModelSpace object
    ms = app.ActiveDocument.ModelSpace
   
    #Add a POINT in ModelSpace
    pt = array.array('d', [0,0,0])
    point = ms.AddPoint(pt)
   
    #Add a LINE in ModelSpace
    pt1 = array.array('d', [1.0,1.0,0])
    pt2 = array.array('d', [2.0,2.0,0])
    line = ms.AddLine(pt1, pt2)
   
    #Add an integer type xdata to the point.
    point.SetXData(array.array("h", [1001, 1070]), ['Test_Application1', 600])
   
    #Add a double type xdata to the line.
    line.SetXData(array.array("h", [1001, 1040]), ['Test_Application2', 132.65])
   
    #Add a string type xdata to the line.
    line.SetXData(array.array("h", [1001, 1000]), ['Test_Application3', 'TestData'])
   
    #Add a list type (a point coordinate in this case) xdata to the line.
    line.SetXData(array.array("h", [1001, 1010]),
	          ['Test_Application4', array.array('d', [2.0,0,0])])
   
    print "Done."


COM events
**********

Some COM objects support events, which allows them to notify the user
of the object when something happens.  The standard COM mechanism is
based on so-called *connection points*.

Note: For the rules that you should observe when implementing event
handlers you should read the implementing_COM_methods_ section in the
|comtypes| server document.

``GetEvents(source, sink, interface=None)``
    This functions connects an event sink to the COM object
    ``source``.

    Events will call methods on the ``sink`` object; the methods must
    be named ``interfacename_methodname`` or ``methodname``.  The
    methods will be called with a ``this`` parameter, plus any
    parameters that the event has.

    ``interface`` is the outgoing interface of the ``source`` object;
    it must be supplied when |comtypes| cannot determine the
    outgoing interface of ``source``.

    ``GetEvents`` returns the advise connection; you should keep the
    connection alive as long as you want to receive events.  To break
    the advise connection simply delete it.

``ShowEvents(source, interface=None)``
    This function contructs an event sink and connects it to the
    ``source`` object for debugging.  The event sink will first print
    out all event names that are found in the outgoing interface, and
    will later print out the events with their arguments as they occur.
    ``ShowEvents`` returns a connection object which must be kept
    alive as long as you want to receive events.  When the object is
    deleted the connection to the source object is closed.

    To actually receive events you may have to call the ``PumpEvents``
    function so that COM works correctly.

``PumpEvents(timeout)``
    This functions runs for a certain time in a way that is required
    for COM to work correctly.  In a single-theaded apartment it runs
    a windows message loop, in a multithreaded apparment it simply
    waits.  The ``timeout`` argument may be a floating point number to
    indicate a time of less than a second.

    Pressing Control-C raises a KeyboardError exception and terminates
    the function immediately.


Examples
++++++++

Here is an example which demonstrates how to find and receive events
from ``stdole.StdFont``:

.. doctest::

    >>> font = CreateObject("StdFont")
    >>> font  # doctest: +ELLIPSIS
    <POINTER(Font) ptr=... at ...>


The ``ShowEvents`` function is a useful helper to get started with the
events of an object in the interactive Python interpreter.

We call ``ShowEvents`` to connect to the events that ``StdFont`` fires.
``ShowEvents`` first lists the events that are present on the
``StdFont`` object:

.. doctest::

    >>> from comtypes.client import ShowEvents
    >>> connection = ShowEvents(font)
    # event found: FontEvents_FontChanged
    >>> connection  # doctest: +ELLIPSIS
    <comtypes.client._events._AdviseConnection object at ...>


We have assigned the return value of the ``ShowEvents`` call to the
variable ``connection``, this variable keeps the connection to ``StdFont``
alive and it will print events as they actually occur.

.. doctest::

    >>> font.Name = 'Arial'
    Event FontEvents_FontChanged(None, 'Name')
    >>> font.Italic = True
    Event FontEvents_FontChanged(None, 'Italic')


The first parameter is always the ``this`` pointer passed as ``None``
for |comtypes|-internal reasons, other parameters depend on the event.

The ``PumpEvents()`` function will run a message loop for a certain time.
|comtypes| prints the events as they are fired with their parameters:

.. doctest::

    >>> from comtypes.client import PumpEvents
    >>> PumpEvents(0.01)  # The output will be in the form of "FontEvents_FontChanged(None, 'Name')".


To terminate the connection, we call the ``disconnect`` method. It may
also be necessary to delete the ``connection`` variable and invoke the
Python garbage collector.  Afterward, no events from ``StdFont`` will
be received anymore.

.. doctest::

    >>> connection.disconnect()
    >>> del connection
    >>> import gc
    >>> _ = gc.collect()
    >>> font.Name = 'Sans'  # Expected nothing


If we want to process the events in our own code, we use the
``GetEvents()`` function in a very similar way.  This function must be
called with the COM object as the first argument, the second parameter
is a Python object, the event sink, that will process the events.  The
event sink should have methods named like the events we want to
process.  It is only required to implement methods for those events
that we want to process, other events are ignored.

The following code defines a class that processes the
``FontEvents_FontChanged`` event, creates an instance of this class
and passes it as second parameter to the ``GetEvents()`` function:

.. doctest::

   >>> from comtypes.client import GetEvents
   >>> class EventSink(object):
   ...     def FontEvents_FontChanged(self, this, PropertyName):
   ...         print("FontChanged", PropertyName)
   ...         # add your code here
   ...
   >>> sink = EventSink()
   >>> connection = GetEvents(font, sink)
   >>> font.Name = 'Arial'
   FontChanged Name


Note that event handler methods support the same calling convention as
COM method implementations in |comtypes|.  So the remarks about
implementing_COM_methods_ should be observed.

Typelibraries
*************

Accessing type libraries
++++++++++++++++++++++++

|comtypes| uses early binding even to custom COM interfaces.  A Python
class, derived from the ``comtypes.IUnknown`` class must be written.
This class describes the interface methods and properties in a way
that is somewhat similar to IDL notation.

It should be possible to write the interface classes manually,
fortunately |comtypes| includes a code generator that does create
modules containing the Python interface class (and more) automatically
from COM typelibraries.

``GetModule(tlib)``
    This function generates Python wrappers for a COM typelibrary.
    When a COM object exposes its own typeinfo, this function is
    called automatically when the object is created.

    ``tlib`` can be the following:

    - an ``ITypeLib`` COM pointer from a loaded typelibrary
    - the pathname of a file containing a type library (``.tlb``,
      ``.exe`` or ``.dll``)
    - a tuple or list containing the typelibrary's GUID, optionally
      along with a major and a minor version numbers if versioning
      is required, plus optionally a LCID.
    - any object that has a ``_reg_libid_`` and ``_reg_version_``
      attributes specifying a type library.

    ``GetModule(tlib)`` generates two Python modules (if not already
    present): a first wrapper module and a second friendly module,
    within the ``comtypes.gen`` package with a single call and
    returns the second friendly module.  If modules are already
    present, it imports the two modules and returns the friendly
    module.

    A first wrapper module is created from the typelibrary, is
    containing interface classes, coclasses, constants, and
    structures.  The module name is derived from the typelibrary
    guid, version numbers and lcid.  The module name is a valid
    Python module name, so it can be imported with an import
    statement.

    A second friendly module is also created in the ``comtypes.gen``
    package with a shorter name that is derived from the type
    library *name* itself.  It does import the wrapper module with an
    abstracted alias ``__wrapper_module__``, also imports interface
    classes, coclasses, constants, and structures from the wrapper
    module, and defines enumerations from typeinfo of the typelibrary
    using `enum.IntFlag`_.  The friendly module can be imported
    easier than the wrapper module because the module name is easier
    to type and read.

    For example, the typelibrary for Scripting Runtime has the name
    ``Scripting`` (this is the name specified in the type library
    IDL file, it is not the filename), the guid is
    ``{420B2830-E718-11CF-893D-00A0C9054228}``, and the version
    number ``1.0``.  The name of the first typelib wrapper module is
    ``comtypes.gen._420B2830_E718_11CF_893D_00A0C9054228_0_1_0`` and
    the name of the second friendly module is ``comtypes.gen.Scripting``.

    When you want to freeze your script with ``py2exe`` you can ensure
    that ``py2exe`` includes these typelib wrappers by writing:

    .. sourcecode:: python

        import comtypes.gen.Scripting

    somewhere.

    *Added in version 1.3.0*: The friendly module imports the wrapper
    module with an abstracted alias ``__wrapper_module__``.

    *Changed in version 1.4.0*: The friendly module defines
    enumerations from typeinfo of the typelibrary.
    Prior to this, the friendly module imported everything from the
    wrapper module, and all names used in enumerations were aliases
    for ``ctypes.c_int``.  Even after version 1.4.0, by modifying the
    codebase as follows, these names can continue to be used as
    aliases for ``c_int`` rather than as enumerations.

    .. sourcecode:: diff

        - from comtypes.gen.friendlymodule import TheName
        + from ctypes import c_int as TheName

    .. sourcecode:: diff

        from comtypes.gen import friendlymodule
        - c_int_alias = friendlymodule.TheName
        + c_int_alias = friendlymodule.__wrapper_module__.TheName

    .. sourcecode:: diff

        - from comtypes.gen import friendlymodule as mod
        + from comtypes.gen.friendlymodule import __wrapper_module__ as mod
        c_int_alias = mod.TheName

``gen_dir``
    This variable determines the directory where the typelib wrappers
    are written to.  If it is ``None``, modules are only generated in
    memory.

    ``comtypes.client.gen_dir`` is calculated when the
    ``comtypes.client`` module is first imported.  It is set to the
    directory of the ``comtypes.gen`` package when this is a valid
    file system path; otherwise it is set to ``None``.

    In a script frozen with ``py2exe`` the directory of ``comtypes.gen``
    is somewhere in a zip-archive, ``gen_dir`` is ``None``, and even
    if tyelib wrappers are created at runtime no attempt is made to
    write them to the file system.  Instead, the modules are generated
    only in memory.

    ``comtypes.client.gen_dir`` can also be set to ``None`` to prevent
    writing typelib wrappers to the file system.  The downside is that
    for large type libraries the code generation can take some time.

Examples
++++++++

Here are several ways to generate the typelib wrapper module for
Scripting Dictionary with the ``GetModule`` function:

.. doctest::

    >>> from comtypes.client import GetModule
    >>> GetModule('scrrun.dll')  # doctest: +ELLIPSIS
    <module 'comtypes.gen.Scripting'...>
    >>> GetModule(('{420B2830-E718-11CF-893D-00A0C9054228}', 1, 0))  # doctest: +ELLIPSIS
    <module 'comtypes.gen.Scripting'...>

Members such as the first wrapper module, interface classes,
coclasses, constants, and enumerations can be referenced from the
friendly module generated by calling the ``GetModule`` function:

.. doctest::

    >>> Scripting = GetModule('scrrun.dll')
    >>> Scripting.__wrapper_module__  # the first wrapper module  # doctest: +ELLIPSIS
    <module 'comtypes.gen._420B2830_E718_11CF_893D_00A0C9054228_0_1_0'...>
    >>> Scripting.IDictionary  # an interface class
    <class 'comtypes.gen._420B2830_E718_11CF_893D_00A0C9054228_0_1_0.IDictionary'>
    >>> Scripting.Dictionary  # a coclass
    <class 'comtypes.gen._420B2830_E718_11CF_893D_00A0C9054228_0_1_0.Dictionary'>
    >>> Scripting.BinaryCompare  # a constant
    0
    >>> Scripting.CompareMethod  # an enumeration
    <flag 'CompareMethod'>
    >>> Scripting.CompareMethod.BinaryCompare  # a member of the enumeration     
    <CompareMethod.BinaryCompare: 0>


This code snippet could be used to generate the typelib wrapper module
for Scripting Dictionary automatically when your script is run, and
would include the module into the exe-file when the script is frozen
by ``py2exe``:

.. doctest::

    >>> import sys
    >>> if not hasattr(sys, 'frozen'):  # doctest: +ELLIPSIS
    ...     from comtypes.client import GetModule
    ...     GetModule('scrrun.dll')
    ...
    <module 'comtypes.gen.Scripting'...>
    >>> import comtypes.gen.Scripting


Other stuff
***********

XXX describe logging, gen_dir, wrap, _manage (?)


.. |comtypes| replace:: ``comtypes``

.. _`WMI monikers`: http://www.microsoft.com/technet/scriptcenter/guide/sas_wmi_jgfx.mspx?mfr=true

.. _`enum.IntFlag`: https://docs.python.org/3/library/enum.html#enum.IntFlag

.. _implementing_COM_methods: server.html#implementing-com-methods
