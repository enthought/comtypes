# comtypes

![Works on Windows only](https://img.shields.io/badge/-Windows-0078D6.svg?logo=windows&style=flat)  
[![PyPI version](https://badge.fury.io/py/comtypes.svg)](https://pypi.org/project/comtypes/) [![PyPI - Python Version](https://img.shields.io/pypi/pyversions/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - License](https://img.shields.io/pypi/l/comtypes)](https://pypi.org/project/comtypes/) [![PyPI - Downloads](https://img.shields.io/pypi/dm/comtypes)](https://pypi.org/project/comtypes/)  
[![GitHub Repo stars](https://img.shields.io/github/stars/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/stargazers) [![GitHub forks](https://img.shields.io/github/forks/enthought/comtypes?style=social)](https://github.com/enthought/comtypes/network/members)  
[![Tidelift Subscription](https://tidelift.com/badges/package/pypi/comtypes)](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=readme)


`comtypes` is a lightweight pure Python [COM](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) package based on the [`ctypes`](https://docs.python.org/library/ctypes.html) foreign function interface library.

`comtypes` allows you to define, call, and implement custom and dispatch-based COM interfaces in pure Python.

`comtypes` requires Windows and Python 3.7 or later.
- Version [1.2.1](https://pypi.org/project/comtypes/1.2.1/) is the last version to support Python 2.7 and 3.3–3.6.
- `comtypes` does not work with Python 3.7.6 and 3.8.1 as reported in [GH-202](https://github.com/enthought/comtypes/issues/202). This bug has been fixed in Python >= 3.7.7 and >= 3.8.2.
- Certain `comtypes` functions may not work correctly in Python 3.8 and 3.9 as reported in [GH-212](https://github.com/enthought/comtypes/issues/212). This bug has been fixed in Python >= 3.10.10 and >= 3.11.2.

## Installation

`comtypes` is available on [PyPI](https://pypi.org/project/comtypes) and can be installed with `pip`:
```sh
# PyPI
pip install comtypes
```

The source code is currently hosted [here](https://github.com/enthought/comtypes) on GitHub.

### Dependencies

`comtypes` is a pure Python package — it has no additional required dependencies.

Optional functionalities can be enabled by installing:
- `numpy`, in order to process arrays as `numpy`'s `ndarray`.
- `mypy` or other static type checkers, in order to interpret type hints.

None of these packages, however, are required in order to run `comtypes`.

## Community

The [GitHub repository](https://github.com/enthought/comtypes) is used for tracking issues, reporting bugs, and contributing to the codebase and documentation.

## For Enterprise

Available as part of the Tidelift Subscription.

This project and the maintainers of thousands of other packages are working with Tidelift to deliver one enterprise subscription that covers all of the open source you use.

[Learn more](https://tidelift.com/subscription/pkg/pypi-comtypes?utm_source=pypi-comtypes&utm_medium=referral&utm_campaign=github).

## Frequently Asked Questions

### Q: Why does this package not support platforms other than Windows?
**A:** The [Microsoft Component Object Model (COM)](https://learn.microsoft.com/en-us/windows/win32/com/com-technical-overview) is a technology that is unique to Windows and is not supported on other platforms.

[The phrase _"COM is a platform-independent"_ in the MS documentation](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) means that COM maintains compatibility across different versions of Windows, and does NOT imply that it is supported on Linux or Mac.

For as long as COM is not supported outside of Windows, there is no plan to port `comtypes` to other platforms.

### Q: Why does `cannot import name 'COMError' from '_ctypes'` error occur when using this package on platforms other than Windows?
**A:** The [`_ctypes`](https://github.com/python/cpython/blob/main/Modules/_ctypes/_ctypes.c) is part of the internal implementation of the [`ctypes`](https://github.com/python/cpython/blob/main/Lib/ctypes/) standard library that exists for Python on all platforms.
However, `COMError` and COM-related features are only implemented in Python for Windows.

In cross-platform software development, care must be taken to ensure that codebases that depend on `comtypes` do not execute in environments other than Windows.

### Q: Despite a script that depends on `comtypes` having run successfully before, a error (`ImportError`, `NameError`, or `SyntaxError`) is raised now, and the same error occurs again and again.

**A:** Executing `py -m comtypes.clear_cache` and then running the script again might resolve the problem.

When `comtypes.client.GetModule` is called (either directly or indirectly), `comtypes` generates Python module files.  
If Python is forced to terminate or crashes in the middle of file generation, the codebase written to the file becomes partial.  
When Python tries to import this unexecutable partial codebase module, an error occurs.

Executing `py -m comtypes.clear_cache` identifies the directories where the "cache module files" are stored and deletes them.  
After deleting these partial modules and running the script again, `comtypes.client.GetModule` is called and executable modules are generated anew.

However, if the script implementation does not use `comtypes.client.GetModule` or processes generated files, it may not be a solution.

## Documentation

The documentation is currently hosted [here](https://pythonhosted.org/comtypes) on PythonHosted.
