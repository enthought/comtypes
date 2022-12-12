# comtypes

[![PyPI version](https://badge.fury.io/py/comtypes.svg)](https://pypi.org/project/comtypes/) [![PyPI - Python Version](https://img.shields.io/pypi/pyversions/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - License](https://img.shields.io/pypi/l/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - Downloads](https://img.shields.io/pypi/dd/comtypes)](https://pypi.org/project/comtypes/)
[![GitHub Repo stars](https://img.shields.io/github/stars/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/stargazers) [![GitHub forks](https://img.shields.io/github/forks/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/network/members)

## About

`comtypes` is a lightweight `Python` COM package, based on the [`ctypes`](https://docs.python.org/library/ctypes.html) FFI library.

`comtypes` allows to define, call, and implement custom and dispatch-based COM interfaces in pure `Python`.

This package works on only Windows.
- [`comtypes==1.1.7`](https://pypi.org/project/comtypes/1.1.7/) is the last version of supporting Windows CE.

## Where to get it

The source code is currently hosted on GitHub at:
https://github.com/enthought/comtypes

Binary installers for the latest released version are available at the [Python Package Index (PyPI)](https://pypi.org/project/comtypes).

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

## Documentation:

The documentation is currently hosted on pythonhosted at:
https://pythonhosted.org/comtypes
