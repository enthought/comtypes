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

## For Enterprise

Available as part of the Tidelift Subscription.

This project and the maintainers of thousands of other packages are working with Tidelift to deliver one enterprise subscription that covers all of the open source you use.

[Learn more](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=referral&utm_campaign=github).

## Documentation:

The documentation is currently hosted on pythonhosted at:
https://pythonhosted.org/comtypes
