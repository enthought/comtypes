####################
The comtypes package
####################

|comtypes| is a *pure Python* COM package based on the ctypes_ ffi
foreign function library.  |ctypes| is included in Python 2.5 and
later, it is also available for Python 2.4 as separate download.

While the pywin32_ package contains superior client side support
for *dispatch based* COM interfaces, it is not possible to access
*custom* COM interfaces unless they are wrapped in C++-code.

The |comtypes| package makes it easy to access and implement both
custom and dispatch based COM interfaces.

.. contents::


Functionalities
***************

.. toctree::
    :maxdepth: 1

    client
    server
    npsupport
    threading


Links
*****

Kourovtsev, Yaroslav (2008). `"Working with Custom COM Interfaces from Python" <http://www.codeproject.com/KB/COM/python-comtypes-interop.aspx>`_

    This article describes how to use |comtypes| to access a custom
    COM object.


Downloads
*********

The |comtypes| project is hosted on github_. Releases can be downloaded from
the github releases_ section.


.. |comtypes| replace:: ``comtypes``

.. |ctypes| replace:: ``ctypes``

.. _ctypes: https://docs.python.org/3/library/ctypes.html

.. _pywin32: https://pypi.org/project/pywin32/

.. _github: https://github.com/enthought/comtypes

.. _releases: https://github.com/enthought/comtypes/releases
