
class Patch(object):
    """
    Implements a class decorator suitable for patching an existing class with
    a new namespace.

    For example, consider this trivial class (that your code doesn't own):

    >>> class MyClass:
    ...     def __init__(self, param):
    ...         self.param = param

    To add attributes to MyClass, you can use MonkeyPatch:

    >>> @Patch(MyClass)
    ... class JustANamespace:
    ...     def print_param(self):
    ...         print(self.param)
    >>> ob = MyClass('foo')
    >>> ob.print_param()
    foo
    """

    def __init__(self, target):
        self.target = target

    def __call__(self, patches):
        for name, value in vars(patches).items():
            if name in vars(ReferenceEmptyClass):
                continue
            setattr(self.target, name, value)

class ReferenceEmptyClass(object):
    """
    This empty class will serve as a reference for attributes present on
    any class.
    """
