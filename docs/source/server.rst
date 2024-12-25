#########################
COM servers with comtypes
#########################

The |comtypes| package makes it easy to access and implement both
custom and dispatch based COM interfaces.

.. contents::

Implementing a simple COM object
********************************

To implement a COM server object in |comtypes| you need to write a type
library describing the coclass, the interface(s) that the object
implements, and (optional) the event interface that the object
supports.  Also you have to write a Python module that defines a class
which implements the object itself.

We will present a short example here that does actually work.

Define the COM interface
++++++++++++++++++++++++

Start writing an IDL file.  It is a good idea to define ``dual``
interfaces, and only use automation compatible data types.

.. sourcecode:: c

  import "oaidl.idl";
  import "ocidl.idl";

  [
          uuid(xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx),
          dual,
          oleautomation
  ]
  interface IMyInterface : IDispatch {
          HRESULT MyMethod([in] INT a, [in] INT b, [out, retval] INT *presult);
  }

  [
  	uuid(xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)
  ]
  library MyTypeLib
  {
          importlib("stdole2.tlb");
  	
          [uuid(xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)]
  		coclass MyObject {
  		[default] interface IMyInterface;
          };
  };

Please note that you must replace the 'xxxx' placeholders in the
section above with separate GUIDs that you must generate yourself.
You can use Python to generate unique GUIDs by running this in a
windows console:

.. sourcecode:: shell

    C:\> python -c "from comtypes import GUID; print GUID.create_new()"
    {26F87CEB-603A-4FFE-8865-DB67A9E3A308}
    C:\> 

The IDL file should now be compiled with the Microsoft MIDL compiler to a
TLB type library file.


Implement the class
+++++++++++++++++++

Generate and import the wrapper module (which is named after the
``library`` statement in the IDL file), and create a subclass of the
``MyObject`` coclass.

Most required class attributes are already defined in the typelib
wrapper file.  You must at least add attributes for registration that
are not in the type library.

.. sourcecode:: python

  import comtypes
  import comtypes.server.localserver

  from comtypes.client import GetModule
  # generate wrapper code for the type library, this needs
  # to be done only once (but also each time the IDL file changes)
  GetModule("mytypelib.tlb")

  from comtypes.gen.MyTypeLib import MyObject

  class MyObjectImpl(MyObject):
      # registry entries
      _reg_threading_ = "Both"
      _reg_progid_ = "MyTypeLib.MyObject.1"
      _reg_novers_progid_ = "MyTypeLib.MyObject"
      _reg_desc_ = "Simple COM server for testing"
      _reg_clsctx_ = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER
      _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE

The meaning of the attributes:

    ``_reg_threading_`` must be set to "Both", "Free", or "Apartment".
    It specifies the apartment model in which the server runs.

    ``_reg_progid_`` and ``_reg_novers_progid`` are optional short
    names that can later be used to specify your object, instead of
    the CLSID in type library.  Typically the type library name plus
    the coclass name plus a version number are combined to form the
    progid, and the type library name plus the coclass name are
    combined to form the version independend progid.

    ``_reg_desc_`` is the (optional) name of the coclass.

    The ``_reg_clsctx_`` constant specifies in which contexts the com
    server can operate.

    The optional ``_regcls_`` constant is only used for com objects
    that run in their own process, see the MSDN docs for more info.
    In |comtypes|, several REGCLS values are defined in the
    ``comtyper.server.localserver`` module.

You do not yet implement any methods on the class, because basic
functionality is already present.

Register and run the object for the first time
++++++++++++++++++++++++++++++++++++++++++++++

A COM object must by registered with Windows, and will also be started
at runtime by Windows.  This magic, on the |comtypes| side, is done by
the ``comtypes.server.register.UseCommandLine`` function.  You should
call it in the ``if __name__ == "__main__"`` block of your script,
with the ``MyObjectImpl`` class:

.. sourcecode:: python

    if __name__ == "__main__":
        from comtypes.server.register import UseCommandLine
        UseCommandLine(MyObjectImpl)

You should now run your script with a ``/regserver`` command line
option, this will write information about your object into the Windows
registry:

.. sourcecode:: shell

    C:\> python myserver.py /regserver

If you have the Microsoft ``OLEVIEW`` utility, you can now open the
"All Objects" item, and look for the "Simple COM server for testing"
object.  If everything works well, you can even create an instance of
your COM object by double clicking the entry, and you will see that
the object implements quite some interfaces already.

You can also create an instance of the object with |comtypes|:

.. sourcecode:: pycon

  >>> from comtypes.client import CreateObject
  >>> x = CreateObject("MyTypelib.MyObject")
  >>> print x
  <POINTER(IMyInterface) ptr=0x1216328 at 1216620>
  >>>

Of course calling a method does not yet work since it is not
implemented in the server script:

.. sourcecode:: pycon

  >>> x.MyMethod(a, 2)
  Traceback (most recent call last):
    File "<stdin>", line 1, in <module>
  _ctypes.COMError: COMError(0x80004001, 'Nicht implementiert', (None, None, None, 0, None))
  >>>

Implementing COM methods
++++++++++++++++++++++++

NOTE: The documentation in this section is also valid for writing
COM event handlers!

In the IDL file, the method signature is defined like this:

.. sourcecode:: c

    HRESULT MyMethod([in] INT a, [in] INT b, [out, retval] INT *presult);

So, this method takes two integers and returns a third one, writing
the latter into a pointer.

You must add e Python method to the class ``MyObject`` that implements
this behaviour.

Determining the method name
---------------------------

The method implementing the ``IMyInterface.MyMethod`` can either be
named ``IMyInterface_MyMethod`` or ``MyMethod``.  Choose a name that
does not conflict with other methods of the class, and that serves
your personal naming conventions.

In |comtypes|, there are two ways to implement COM server methods.
You can choose between a 'low level' and a 'high level' implementation
strategy, on a method by method basis (the names 'Low level' and 'high
level' are probably misleading a bit, suggestions for better names
would be welcomed).  |comtypes| uses different calling conventions for
'low level' and 'high level' method implementations.

|comtypes| inspects the method for the name of the second parameter,
just after the ``self`` parameter:

  **If the second parameter is present and is named ``this`` then the
  low level calling convention is used.  If the second parameter is
  not present, or is not named ``this``, then the high level calling
  convention is used.**


Low level implementation
------------------------

A low-level method implementation is called with the following arguments:

- the usual ``self`` argument

- for the ``this`` argument either ``None`` is passed, or the address
  of the COM object itself as an integer.  The value of it can usually
  and should be ignored.

- any other arguments listed in the IDL method signature.

[in] parameters from the method signature are usually converted to
native Python objects, if possible.  For [out] or [out, retval]
parameters ctypes pointer instances are passed, you are required to
put the result value into the pointer(s).

A low level method implementation must return a numerical HRESULT
value, which specifies a success or failure code for the operation.
The usual ``S_OK`` success code has a value of zero, but for
convenience you can also return None instead.

So, a sample low-level implementation for ``MyMethod`` for our object
would be this, assuming we want to return the sum of the two [in]
parameters:

.. sourcecode:: python

  ...
  class MyObjectImpl(MyObject):
  ...
      # Note the 'this' second parameter
      def MyMethod(self, this, a, b, presult):
          presult[0] = a + b
          return 0

High level implementation
-------------------------

A high-level method implementation is called with the following parameters:

- the usual ``self`` argument

- the [in] parameters from the IDL method signature.

If there is a single [out] or [out, retval] parameter, then the method
must return the result value; if there are more than one [out] or
[out, retval] parameters, then a tuple containing the correct number
must be returned.  If there are no [out] or [out, retval] parameters,
the return value does not matter and is ignored.

A sample high-level implementation for ``MyMethod`` is this:

.. sourcecode:: python

  class MyObjectImpl(MyObject):
  ...
      # Note: NO second 'this' parameter
      def MyMethod(self, a, b):
          return a + b

Choosing between low-level or high-level implementation
-------------------------------------------------------

Both implementation strategies have their own advantages and
disadvantages, so you should choose between them on a case by case
basis:

Low-level makes it easy to return special HRESULT values in the case
that your object requires it.

High-level is usually easier to write, and is compatible with the
normal calling convention that Python also chooses.  However, it is
more difficult to specify the HRESULT value to return in case you want
to communicate error codes to the caller.

Run the object again and test the method
++++++++++++++++++++++++++++++++++++++++

We can now create the object and test the implemented method:

.. sourcecode:: pycon

  >>> from comtypes.client import CreateObject
  >>> x = CreateObject("MyTypelib.MyObject")
  >>> print x
  <POINTER(IMyInterface) ptr=0x1216328 at 1216620>
  >>> print x.MyMethod(42, 5)
  47
  >>>

More details on COM objects
***************************

To be written...

.. |comtypes| replace:: ``comtypes``
