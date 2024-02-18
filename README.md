# comtypes

![Works on Windows only](https://img.shields.io/badge/-Windows-0078D6.svg?logo=windows&style=flat)  
[![PyPI version](https://badge.fury.io/py/comtypes.svg)](https://pypi.org/project/comtypes/) [![PyPI - Python Version](https://img.shields.io/pypi/pyversions/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - License](https://img.shields.io/pypi/l/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - Downloads](https://img.shields.io/pypi/dm/comtypes)](https://pypi.org/project/comtypes/)  
[![GitHub Repo stars](https://img.shields.io/github/stars/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/stargazers) [![GitHub forks](https://img.shields.io/github/forks/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/network/members)  
[![Tidelift Subscription](https://tidelift.com/badges/package/pypi/comtypes)](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=readme)

## About

`comtypes` is a lightweight `Python` [COM](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) package, based on the [`ctypes`](https://docs.python.org/library/ctypes.html) FFI library.

`comtypes` allows to define, call, and implement custom and dispatch-based COM interfaces in pure `Python`.

This package works on Windows only.
- [`comtypes==1.1.7`](https://pypi.org/project/comtypes/1.1.7/) is the last version supporting Windows CE.

Available on `Python` 3.7-3.12.
- [`comtypes==1.2.1`](https://pypi.org/project/comtypes/1.2.1/) is the last version supporting `Python` 2.7 and 3.3-3.6.
- In `Python` 3.7.6 and 3.8.1, `comtypes` would not work as reported in [GH-202](https://github.com/enthought/comtypes/issues/202).  
This bug has been fixed in `Python` >= 3.7.7 and >= 3.8.2.
- In `Python` 3.8 and 3.9, some of `comtypes` functionalities may not work correctly as reported in [GH-212](https://github.com/enthought/comtypes/issues/212).  
This bug has been fixed in `Python` >= 3.10.10 and >= 3.11.2.

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
#### In friendly modules, the names that were used as aliases for `ctypes.c_int` will be used for enumeration types implemented with `enum`
`comtypes.client.GetModule` generates two `Python` modules in the `comtypes.gen` package with a single call.

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
