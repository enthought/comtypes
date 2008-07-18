####################
The comtypes package
####################

|comtypes| is a *pure Python* COM package based on the ctypes_ ffi
foreign function library.  **ctypes** is included in Python 2.5 and
later, it is also available for Python 2.4 as separate download.

While the **pywin32** package contains superior client side support
for *dispatch based* COM interfaces, it is not possible to access
*custom* COM interfaces unless they are wrapped in C++-code.

The |comtypes| package makes it easy to access and implement both
custom and dispatch based COM interfaces.

This document describes |comtypes| version 0.4.1.

.. contents::

The comtypes.client package
***************************

The **comtypes.client** package implements the high-level |comtypes|
functionality.

Creating and accessing COM objects
++++++++++++++++++++++++++++++++++

**comtypes.client** exposes three functions that allow to create or
access COM objects.

``CreateObject(progid, clsctx=None, machine=None, interface=None)``
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

``CoGetObject(displayname, interface=None)``
    Create a named COM object and returns an interface pointer to it.
    For the interpretation of ``displayname`` consult the Microsoft
    documentation for the Windows ``CoGetObject`` function.
    ``"winmgmts:"``, for example, is the displayname for `WMI
    monikers`_:

    .. sourcecode:: python

        wmi = CoGetObject("winmgmts:")

    ``interface`` has the same meaning as in the ``CreateObject``
    function.

``GetActiveObject(progid, interface=None)``
    Returns a pointer to an already running object.  ``progid``
    specifies the active object from the OLE registration database.

    The ``GetActiveObject`` function succeeds when the COM object is
    already running, and has registered itself in the COM running
    object table.  Not all COM objects do this.

All the three functions mentioned above will create the typelib
wrapper automatically if the object provides type information.  If the
type library is not exposed by the object itself, the wrapper can be
created by calling the ``GetModule`` function.


Using COM objects
+++++++++++++++++

The COM interface pointer that is returned by one of the creation
functions (``CreateObject``, ``CoGetObject``, or ``GetActiveObject``)
exposes methods and properties of the interface.

Since ``comtypes`` uses early binding to COM interfaces (when type
information is exposed by the COM object), the interface methods and
properties are available for introspection.  The Python builtin
``help`` function can be used to get an overview of them.

``MSScriptControl.ScriptControl`` is the progid of the MS scripting
engine; this is an interesting COM object that allows to execute
JScript or VBScript programs.  Here_ is the complete output of these
commands:

.. sourcecode:: pycon

    >>> from comtypes.client import CreateObject
    >>> engine = CreateObject("MSScriptControl.ScriptControl")
    >>> help(engine)
    .....
    >>>

.. _Here: scriptcontrol.html


Calling methods
---------------

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
--------------------

COM properties present some challenges.  Properties can be read-write,
read-only, or write-only.  They may have zero, one, or more arguments;
arguments may even be optional.

Properties without arguments can be accessed in the usual way.  This
example demonstrates the ``Visible`` property of Internet Explorer:

.. sourcecode:: pycon

    >>> ie = CreateObject("InternetExplorer.Application")
    >>> print ie.Visible
    False
    >>> ie.Visible = True
    >>>


Properties with arguments (named properties)
............................................

Properties with arguments can be accessed using index notation.
The following example starts Excel, creates a new workbook, and
accesses the contents of some cells in the ``xlRangeValueDefault``
format (this code has been tested with Office 2003):

.. sourcecode:: pycon

    >>> xl = CreateObject("Excel.Application")
    >>> xl.Workbooks.Add()
    >>> from comtypes.gen.Excel import xlRangeValueDefault
    >>> xl.Range["A1", "C1"].Value[xlRangeValueDefault] = (10,"20",31.4)
    >>> print xl.Range["A1", "C1"].Value[xlRangeValueDefault]
    (10, "20", 31.4)
    >>>


Properties with optional arguments
..................................

If you look into the Excel type library (or the generated
*comtypes.gen* wrapper module) you will find that the parameter for
the ``.Value`` property is optional, so it would be possible to get or
set this property without the need to pass (or even know) the
``xlRangeValueDefault`` argument.

Unfortunately, Python does not allow indexing without arguments:

.. sourcecode:: pycon

    >>> xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
      File "<stdin>", line 1
        xl.Range["A1", "C1"].Value[] = (10,"20",31.4)
                                   ^
    SyntaxError: invalid syntax
    >>> print xl.Range["A1", "C1"].Value[]
      File "<stdin>", line 1
        print xl.Range["A1", "C1"].Value[]
                                         ^
    SyntaxError: invalid syntax
    >>>

So, |comtypes| must provide some ways to access these properties.  To
*get* a named property without passing any argument, you can *call*
the property:

.. sourcecode:: pycon

    >>> print xl.Range["A1", "C1"].Value()
    (10, "20", 31.4)
    >>>

It is also possible to index with an empty tuple:

.. sourcecode:: pycon

    >>> print xl.Range["A1", "C1"].Value[()]
    (10, "20", 31.4)
    >>>

To *set* a named property without passing any argument, you can
also use the empty tuple index trick:

.. sourcecode:: pycon

    >>> xl.Range["A1", "C1"].Value[()] = (1, 2, 3)
    >>>

.. This is not (yet?) implemented.  Would is be useful?
   Another way is to assing to the tuple in the normal way:

      >>> xl.Range["A1", "C1"].Value = (1, 2, 3)
      >>>

The lcid parameter
------------------

Some COM methods or properties have an optional ``lcid`` parameter.
This parameter is used to specify a langauge identifier.  The
generated modules always pass 0 (zero) for this parameter.  If this is
not what you want you have to edit the generated code.

Converting data types
---------------------

|comtypes| usually converts arguments and results between COM and
Python in just the way one would expect.

``VARIANT`` parameters sometimes requires special care.  A ``VARIANT``
can hold a lot of different types - simple ones like integers, floats,
or strings, also more complicated ones like single dimensional or even
multidimensional arrays.  The value a ``VARIANT`` contains is
specified by a *typecode* that comtypes automatically assigns.

When you pass simple sequences (lists or tuples) as VARIANT
parameters, the COM server will receive a VARIANT containing a
SAFEARRAY of VARIANTs with the typecode ``VT_ARRAY | VT_VARIANT``.

Some COM server methods, however, do not accept such arrays, they
require for example an array of shorT integers with the typecode
``VT_ARRAY | VT_I2``, an array of integers with typecode ``VT_ARRAY |
VT_INT``, or an array a strings with typecode ``VT_ARRAY | VT_BSTR``.

To create these variants you must pass an instance of the Python
``array.array`` with the correct Python typecode to the COM method.
The mapping of the ``array.array`` typecode to the ``VARIANT``
typecode is defined in the comtypes.automation module by a
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

XXX Add some simple examples

COM events
++++++++++

Some COM objects support events, which allows them to notify the user
of the object when something happens.  The standard COM mechanism is
based on so-called *connection points*.

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
--------

Here is an example which demonstrates how to find and receive events
from Excel:

.. sourcecode:: pycon

    >>> from comtypes.client import CreateObject
    >>> xl = CreateObject("Excel.Application")
    >>> xl.Visible = True
    >>> print xl
    <POINTER(_Application) ptr=0x29073c at c156c0>
    >>> 

The ``ShowEvents`` function is a useful helper to get started with the
events of an object in the interactive Python interpreter.

We call ``ShowEvents`` to connect to the events that Excel fires.
``ShowEvents`` first lists the events that are present on the
``_Application`` object:

.. sourcecode:: pycon

   >>> from comtypes.client import ShowEvents
   >>> connection = ShowEvents(xl)
   # event found: AppEvents_WorkbookSync
   # event found: AppEvents_WindowResize
   # event found: AppEvents_WindowActivate
   # event found: AppEvents_WindowDeactivate
   # event found: AppEvents_SheetSelectionChange
   # event found: AppEvents_SheetBeforeDoubleClick
   # event found: AppEvents_SheetBeforeRightClick
   # event found: AppEvents_SheetActivate
   # event found: AppEvents_SheetDeactivate
   # event found: AppEvents_SheetCalculate
   # event found: AppEvents_SheetChange
   # event found: AppEvents_NewWorkbook
   # event found: AppEvents_WorkbookOpen
   # event found: AppEvents_WorkbookActivate
   # event found: AppEvents_WorkbookDeactivate
   # event found: AppEvents_WorkbookBeforeClose
   # event found: AppEvents_WorkbookBeforeSave
   # event found: AppEvents_WorkbookBeforePrint
   # event found: AppEvents_WorkbookNewSheet
   # event found: AppEvents_WorkbookAddinInstall
   # event found: AppEvents_WorkbookAddinUninstall
   # event found: AppEvents_SheetFollowHyperlink
   # event found: AppEvents_SheetPivotTableUpdate
   # event found: AppEvents_WorkbookPivotTableCloseConnection
   # event found: AppEvents_WorkbookPivotTableOpenConnection
   # event found: AppEvents_WorkbookBeforeXmlImport
   # event found: AppEvents_WorkbookAfterXmlImport
   # event found: AppEvents_WorkbookBeforeXmlExport
   # event found: AppEvents_WorkbookAfterXmlExport
   >>> print connection
   <comtypes.client._events._AdviseConnection object at 0x00C16AD0>
   >>>

We have assigned the return value of the ``ShowEvents`` call to the
variable ``connection``, this variable keeps the connection to Excel
alive and it will print events as they actually occur.

To receive COM events correctly, it is important to run a message
loop; the ``PumpEvents()`` function will do that for a certain time.
Here is what happens when we call this function and in the meantime
interactively open an Excel worksheet.  ``comtypes`` prints the events
as they are fired with their parameters:

.. sourcecode:: pycon

   >>> from comtypes.client import PumpEvents
   >>> PumpEvents(30)
   Event AppEvents_WorkbookOpen(None, <POINTER(_Workbook) ptr=...>)
   Event AppEvents_WorkbookActivate(None, <POINTER(_Workbook) ptr=...>)
   Event AppEvents_WindowActivate(None, <POINTER(Window) ptr=...>, <POINTER(_Workbook) ptr=...>)
   >>>

The first parameter is always the ``this`` pointer passed as ``None``
for comtypes-internal reasons, other parameters depend on the event.
To terminate the connection we simply delete the ``connection``
variable; it may be required to call the Python garbage collector to
terminate the connection immediately, and we will not receive any
events from Excel anymore:

.. sourcecode:: pycon

   >>> del connection
   >>> import gc; gc.collect()
   123
   >>>

If we want to process the events in our own code, we use the
``GetEvents()`` function in a very similar way.  This function must be
called with the COM object as the first argument, the second parameter
is a Python object, the event sink, that will process the events.  The
event sink should have methods named like the events we want to
process.  It is only required to implement methods for those events
that we want to process, other events are ignored.

The following code defines a class that processes the
``AppEvents_WorkbookOpen`` event, creates an instance of this class
and passes it as second parameter to the ``GetEvents()`` function:

.. sourcecode:: pycon

   >>> from comtypes.client import GetEvents
   >>> class EventSink(object):
   ...     def AppEvents_WorkbookOpen(self, this, workbook):
   ...         print "WorkbookOpened", workbook
   ...         # add your code here
   ...
   >>> sink = EventSink()
   >>> connection = GetEvents(xl, sink)
   >>> PumpEvents(30)
   WorkbookOpened <POINTER(_Workbook) ptr=0x291944 at 1853120>
   >>>


Typelibraries
+++++++++++++

Accessing type libraries
------------------------

|comtypes| uses early binding even to custom COM interfaces.  A Python
class, derived from the ``comtypes.IUnknown`` class must be written.
This class describes the interface methods and properties in a way
that is somewhat similar to IDL notation.

It should be possible to write the interface classes manually,
fortunately |comtypes| includes a code generator that does create
modules containing the Python interface class (and more) automatically
from COM typelibraries.

``GetModule(tlib)``

    This function generates a Python wrapper for a COM typelibrary.
    When a COM object exposes its own typeinfo, this function is
    called automatically when the object is created.

    ``tlib`` can be an **ITypeLib** COM pointer from a loaded
    typelibrary, the pathname of a file containing a type library
    (.tlb, .exe or .dll), a tuple or list containing the GUID of a
    typelibrary, a major and a minor version number, plus optionally a
    LCID, or any object that has a _reg_libid_ and _reg_version_
    attributes specifying a type library.

    ``GetModule(tlib)`` generates a Python module (if not already
    present) from the typelibrary, containing interface classes,
    coclasses, constants, and structures and returns the module object
    itself.  The modules are generated inside the ``comtypes.gen``
    package.  The module name is derived from the typelibrary guid,
    version number and lcid.  The module name is a valid Python module
    name, so it can be imported with an import statement.  A second
    wrapper module is also created in the comtypes.gen package with a
    shorter name that is derived from the type library *name* itself,
    this does import everything from the real wrapper module but can
    be imported easier because the module name is easier to type.

    For example, the typelibrary for Internet Explorer has the name
    ``SHDocVw`` (this is the name specified in the type library IDL
    file, it is not the filename), the guid is
    ``{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}``, and the version number
    ``1.1``.  The name of the real typelib wrapper module is
    ``comtypes.gen._EAB22AC0_30C1_11CF_A7EB_0000C05BAE0B_0_1_1`` and
    the name of the second wrapper is ``comtypes.gen.SHDocVw``.

    When you want to freeze your script with py2exe you can ensure
    that py2exe includes these typelib wrappers by writing:

    .. sourcecode:: python

        import comtypes.gen.SHDocVw

    somewhere.

``gen_dir``

    This variable determines the directory where the typelib wrappers
    are written to.  If it is ``None``, modules are only generated in
    memory.

    ``comtypes.client.gen_dir`` is calculated when the
    **comtypes.client** module is first imported.  It is set to the
    directory of the **comtypes.gen** package when this is a valid
    file system path; otherwise it is set to ``None``.

    In a script frozen with py2exe the directory of **comtypes.gen**
    is somewhere in a zip-archive, ``gen_dir`` is ``None``, and even
    if tyelib wrappers are created at runtime no attempt is made to
    write them to the file system.  Instead, the modules are generated
    only in memory.

    ``comtypes.client.gen_dir`` can also be set to ``None`` to prevent
    writing typelib wrappers to the file system.  The downside is that
    for large type libraries the code generation can take some time.

Examples
--------

Here   are several ways   to generate the  typelib  wrapper module for
Internet Explorer with the ``GetModule`` function:

.. sourcecode:: pycon

   >>> from comtypes.client import GetModule
   >>> GetModule("shdocvw.dll")
   >>> GetModule(["{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}", 1, 1)
   >>>

This code snippet could be used to generate the typelib wrapper module
for Internet Explorer automatically when your script is run, and would
include the module into the exe-file when the script is frozen by
py2exe:

.. sourcecode:: pycon

    >>> import sys
    >>> if not hasattr(sys, "frozen"):
    >>>     from comtypes.client import GetModule
    >>>     GetModule("shdocvw.dll")
    >>> import comtypes.gen.ShDocVw
    >>>


Case sensitivity
----------------

In principle, COM is a case insensitive technology (probably because
of Visual Basic).  Type libraries generated from IDL files, however,
do *not* always even preserve the case of identifiers; see for example
http://support.microsoft.com/kb/220137.

Python (and C/C++) are case sensitive languages, so |comtypes| is also
case sensitive.  This means that you have to call
``obj.QueryInterface(...)``, it will not work to write
``obj.queryinterface(...)``.

To work around the problems that you get when the case of identifiers
in the type library (and in the generated Python module for this
library) is not the same as in the IDL file, |comtypes| allows to have
case insensitive attribute access for methods and properties in COM
interfaces.  This behaviour is enabled by setting the
``_case_insensitive_`` attribute of a Python COM interface to
``True``.  In case of derived COM interfaces, case sensitivity is
enabled or disabled separately for each interface.

The code generated by the ``GetModule`` function sets this attribute
to ``True``.  Case insensitive access has a small performance penalty,
if you want to avoid this, you should edit the generated code and set
the ``_case_insensitive_`` attribute to ``False``.


Threading
+++++++++

XXX mention single threaded apartments, multi threaded apartments.
``sys.coinit_flags``, ``CoInitialize``, ``CoUninitialize`` and so on.
All this is pretty advanced stuff.

XXX mention threading issues, message loops

Other stuff
+++++++++++

XXX describe logging, gen_dir, wrap, _manage (?)

Links
+++++

Yaroslav Kourovtsev has written an article_ titled "Working with custom
COM interfaces from Python" that describes how to use |comtypes| to
access a custom COM object.

.. _article:  http://www.codeproject.com/KB/COM/python-comtypes-interop.aspx

Downloads
*********

Releases can be downloaded in the sourceforge files_ section.

The |comtypes| project is hosted on sourceforge_.

.. include:: footer.rst

.. |comtypes| replace:: **comtypes**

.. _`WMI monikers`: http://www.microsoft.com/technet/scriptcenter/guide/sas_wmi_jgfx.mspx?mfr=true

.. _ctypes: http://starship.python.net/crew/theller/ctypes

.. _sourceforge: http://sourceforge.net/projects/comtypes

.. _files: http://sourceforge.net/project/showfiles.php?group_id=115265