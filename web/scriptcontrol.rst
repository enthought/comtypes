The Python ``help`` function can be used to show the methods and
properties of a COM interface, showing docstrings when they are
available::

    >>> from comtypes.client import CreateObject
    >>> engine = CreateObject("MSScriptControl.ScriptControl")
    >>> help(engine)

.. include:: scriptcontrol.txt
   :literal:

.. include:: footer.rst
