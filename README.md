# comtypes

[![PyPI version](https://badge.fury.io/py/comtypes.svg)](https://pypi.org/project/comtypes/) [![PyPI - Python Version](https://img.shields.io/pypi/pyversions/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - License](https://img.shields.io/pypi/l/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - Downloads](https://img.shields.io/pypi/dd/comtypes)](https://pypi.org/project/comtypes/)
[![GitHub Repo stars](https://img.shields.io/github/stars/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/stargazers) [![GitHub forks](https://img.shields.io/github/forks/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/network/members)

## About

`comtypes` is a lightweight `Python` COM package, based on the [`ctypes`](https://docs.python.org/library/ctypes.html) FFI library.

`comtypes` allows to define, call, and implement custom and dispatch-based COM interfaces in pure `Python`.

This package works on Windows only.
- [`comtypes==1.1.7`](https://pypi.org/project/comtypes/1.1.7/) is the last version supporting Windows CE.

Available on `Python` 2.7 and 3.3-3.11.
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
For the time being, the development target branch of this package will be the [`drop_py2` branch](https://github.com/enthought/comtypes/tree/drop_py2) and the [`master` branch](https://github.com/enthought/comtypes/tree/master) will be in maintenance-only-mode.

As the name suggests, `drop_py2` is a mid-term-planning branch to drop supporting `Python` 2.x from this package, and start supporting `Python` 3.x only.

The `drop_py2` branch will be merged into the `master` branch in the future, and the `master` branch will back to the development target branch and be renamed to `main`.  
Until then, no changes will be made to the `master` branch except in the case of corresponding to regressions.

Please see [the GitHub issue](https://github.com/enthought/comtypes/issues/392) for policy and progress.

## Documentation:

The documentation is currently hosted on pythonhosted at:
https://pythonhosted.org/comtypes
