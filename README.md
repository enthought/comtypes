# comtypes

[![PyPI version](https://badge.fury.io/py/comtypes.svg)](https://pypi.org/project/comtypes/) [![PyPI - Python Version](https://img.shields.io/pypi/pyversions/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - License](https://img.shields.io/pypi/l/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - Downloads](https://img.shields.io/pypi/dd/comtypes)](https://pypi.org/project/comtypes/)
[![GitHub Repo stars](https://img.shields.io/github/stars/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/stargazers) [![GitHub forks](https://img.shields.io/github/forks/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/network/members)
[![Tidelift Subscription](https://tidelift.com/badges/package/pypi/comtypes)](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=readme)

## About

`comtypes` is a lightweight `Python` COM package, based on the [`ctypes`](https://docs.python.org/library/ctypes.html) FFI library.

`comtypes` allows to define, call, and implement custom and dispatch-based COM interfaces in pure `Python`.

This package works on Windows only.
- [`comtypes==1.1.7`](https://pypi.org/project/comtypes/1.1.7/) is the last version supporting Windows CE.

Available on `Python` 2.7 and 3.3-3.12.
- [Supporting `Python` 2.7 will be dropped](#ongoing-plans).

## Where to get it

The source code is currently hosted on GitHub at:
https://github.com/enthought/comtypes

An installer for the latest released version is available at the [Python Package Index (PyPI)](https://pypi.org/project/comtypes).

```sh
# PyPI
pip install comtypes
```

## Dependencies

`comtypes` requires no third-party packages to run - this is truly **pure** `Python` package.

Optional features include the follows...
- to process arrays as `numpy`'s `ndarray`
- type hints be interpreted by `mypy` or several static type checkers

But these third-parties are not required as a prerequisite for runtime.

## Community of the developers

Tracking issues, reporting bugs and contributing to the codebase and documentation are on GitHub at:
https://github.com/enthought/comtypes

<a id="ongoing-plans"></a>
### Ongoing plans
#### Drop supporting `Python` 2.x from this package
For the time being, the development target branch of this package is the [`drop_py2` branch](https://github.com/enthought/comtypes/tree/drop_py2) and the [`main` branch](https://github.com/enthought/comtypes/tree/main) is in maintenance-only-mode.

As the name suggests, `drop_py2` is a mid-term-planning branch to drop supporting `Python` 2.x from this package, and start supporting `Python` 3.x only.

The codebase changes in the `drop_py2` branch will be merged into the `main` branch in the future, and the `main` branch will back to the development target branch.  
Until then, no changes will be made to the `main` branch except in the case of corresponding to regressions.

Please refer to [the GitHub issue](https://github.com/enthought/comtypes/issues/392) for details.

#### `GetModule` will also generate static typing for methods and properties
`comtypes.client.GetModule` generates Python wrapper modules from a COM type library, containing interface classes, coclasses, constants, and structures. The `.py` files are generated in the `comtypes.gen` package.

In the current `comtypes` specification, type checkers could not infer static type information from generated modules codebase, since methods and properties were mostly defined and implemented by metaclasses.  
In future release, in generated modules, static typing will be added to part of methods and properties.

Static type information is added only under [`if TYPE_CHECKING:`](https://docs.python.org/3/library/typing.html#typing.TYPE_CHECKING) blocks. Consequently, it will **not** override any methods defined with metaclasses at runtime, ensuring that the runtime behavior remains unchanged.

Please refer to [the GitHub issue](https://github.com/enthought/comtypes/issues/400) for details.

#### In friendly modules, the names that were used as aliases for `ctypes.c_int` will be used for enumeration types implemented with `enum`
`comtypes.client.GetModule` generates two Python modules in the `comtypes.gen` package with a single call.

A first wrapper module is created with a long name that is derived from the type library guid, version number and lcid. It contains interface classes, coclasses, constants, and structures.  
A second friendly module is created with a shorter name derived from the type library name itself. It imports items from the wrapper module, and will be the module returned from `GetModule`.

In the current `comtypes` specification, if a COM type kind is defined as an enumeration type, that type name is used as an alias for [`ctypes.c_int`](https://docs.python.org/3/library/ctypes.html#ctypes.c_int) in the wrapper module, and that symbol is imported into the friendly module.  
In future release, in friendly modules, their names will no longer be aliases for `c_int`. Instead, they will be defined as enumerations implemented with [`enum`](https://docs.python.org/3/library/enum.html).

When imported into the friendly module, the wrapper module will be aliased with an abstracted name (`__wrapper_module__`). This allows users to continue using the old definitions by modifying the import sections of their codebase.

Please refer to [the GitHub issue](https://github.com/enthought/comtypes/issues/345) for details.

## For Enterprise

Available as part of the Tidelift Subscription.

This project and the maintainers of thousands of other packages are working with Tidelift to deliver one enterprise subscription that covers all of the open source you use.

[Learn more](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=referral&utm_campaign=github).

## Documentation:

The documentation is currently hosted on pythonhosted at:
https://pythonhosted.org/comtypes
