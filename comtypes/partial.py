"""Module for partial classes.

To declare a class partial, inherit from partial.partial and from
the full class, like so

from partial import partial
import original_module

class ExtendedClass(partial, original_module.FullClass):
    def additional_method(self, args):
        body
    more_methods

After this class definition is executed, original_method.FullClass
will have all the additional properties defined in ExtendedClass;
the name ExtendedClass is of no importance (and becomes an alias
for FullClass).
It is an error if the original class already contains the
definitions being added, unless they are methods declared
with @replace.
"""

class _MetaPartial(type):
    "Metaclass implementing the hook for partial class definitions."

    def __new__(cls, name, bases, dict):
        if not bases:
            # It is the class partial itself
            return type.__new__(cls, name, bases, dict)
        if len(bases) != 2:
            raise TypeError("A partial class definition must have only one base class to extend")
        base = bases[1]
        for k, v in dict.items():
            if k == '__module__':
                # Ignore implicit attribute
                continue
            if k in base.__dict__:
                if hasattr(v, '__noreplace'):
                    continue
                if not hasattr(v, '__replace'):
                    raise TypeError("%r already has %s" % (base, k))
            setattr(base, k, v)
        # Return the original class
        return base

class partial:
    "Base class to declare partial classes. See module docstring for details."
    __metaclass__ = _MetaPartial

def replace(f):
    """Method decorator to indicate that a method shall replace
    the method in the full class."""
    f.__replace = True
    return f

def noreplace(f):
    """Method decorator to indicate that a method definition shall
    silently be ignored if it already exists in the full class."""
    f.__noreplace = True
    return f
