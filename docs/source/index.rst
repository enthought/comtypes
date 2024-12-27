########################
The ``comtypes`` package
########################

|comtypes| is a *pure Python* COM package based on the
`ctypes <https://docs.python.org/3/library/ctypes.html>`_ ffi
foreign function library.  |ctypes| is included in Python 2.5 and
later, it is also available for Python 2.4 as separate download.

While the `pywin32 <https://pypi.org/project/pywin32/>`_ package
contains superior client side support for *dispatch based* COM
interfaces, it is not possible to access *custom* COM interfaces
unless they are wrapped in C++-code.

The |comtypes| package makes it easy to access and implement both
custom and dispatch based COM interfaces.

.. contents::


Functionalities
***************

.. toctree::
    :maxdepth: 1

    client
    server
    com_interfaces
    npsupport
    threading


Links
*****

Kourovtsev, Yaroslav (2008). `"Working with Custom COM Interfaces from Python" <http://www.codeproject.com/KB/COM/python-comtypes-interop.aspx>`_

    This article describes how to use |comtypes| to access a custom
    COM object.

Chen, Alicia (2012). `"Comtypes: How Dropbox learned to stop worrying and love the COM" <https://dropbox.tech/infrastructure/adventures-with-comtypes>`_

    This article describes Dropbox's experience using |comtypes| to
    interact with COM objects in a Windows application.


Downloads
*********

The |comtypes| project is hosted on `GitHub <https://github.com/enthought/comtypes>`_.
Releases can be downloaded from the `GitHub releases <https://github.com/enthought/comtypes/releases>`_
section.


Installation
************

|comtypes| is available on `PyPI <https://pypi.org/project/comtypes>`_ and
can be installed with ``pip``:

.. sourcecode:: shell

    pip install comtypes


.. |comtypes| replace:: ``comtypes``

.. |ctypes| replace:: ``ctypes``
