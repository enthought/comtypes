from __future__ import print_function
import ctypes
import importlib
import logging
import os
import sys
import types
if sys.version_info >= (3, 0):
    base_text_type = str
    import winreg
else:
    base_text_type = basestring
    import _winreg as winreg

from comtypes import GUID, TYPE_CHECKING, typeinfo
import comtypes.client
from comtypes.tools import codegenerator, tlbparser

if TYPE_CHECKING:
    from typing import Any, Tuple, List, Optional, Dict, Union as _UnionT


logger = logging.getLogger(__name__)

PATH = os.environ["PATH"].split(os.pathsep)


def _my_import(fullname):
    # type: (str) -> types.ModuleType
    """helper function to import dotted modules"""
    import comtypes.gen as g
    if comtypes.client.gen_dir and comtypes.client.gen_dir not in g.__path__:
        g.__path__.append(comtypes.client.gen_dir)  # type: ignore
    return importlib.import_module(fullname)


def _resolve_filename(tlib_string, dirpath):
    # type: (str, str) -> Tuple[str, bool]
    """Tries to make sense of a type library specified as a string.

    Args:
        tlib_string: type library designator
        dirpath: a directory to relativize the location

    Returns:
        (abspath, True) or (relpath, False):
            where relpath is an unresolved path.
    """
    assert isinstance(tlib_string, base_text_type)
    # pathname of type library
    if os.path.isabs(tlib_string):
        # a specific location
        return tlib_string, True
    elif dirpath:
        abspath = os.path.normpath(os.path.join(dirpath, tlib_string))
        if os.path.exists(abspath):
            return abspath, True
    # try with respect to cwd (if _getmodule executed from command line)
    abspath = os.path.abspath(tlib_string)
    if os.path.exists(abspath):
        return abspath, True
    # Otherwise it may still be that the file is on Windows search
    # path for typelibs, and we leave the pathname alone.
    return tlib_string, False


def GetModule(tlib):
    # type: (_UnionT[Any, typeinfo.ITypeLib]) -> types.ModuleType
    """Create a module wrapping a COM typelibrary on demand.

    'tlib' must be ...
    - an `ITypeLib` COM pointer instance
    - an absolute pathname of a type library
    - a relative pathname of a type library
      - interpreted as relative to the callers `__file__`, if this exists
    - a COM CLSID `GUID`
    - a `tuple`/`list` specifying the typelib
      - `List[_UnionT[str, int]]`
      - `(libid: str[, wMajorVerNum: int, wMinorVerNum: int[, lcid: int]])`
    - an object with `_reg_libid_: str` and `_reg_version_: Iterable[int]`

    This function determines the module name from the typelib
    attributes, then tries to import it.  If that fails because the
    module doesn't exist, the module is generated into the
    `comtypes.gen` package.

    It is possible to delete the whole `comtypes/gen` directory to
    remove all generated modules, the directory and the `__init__.py`
    file in it will be recreated when needed.

    If `comtypes.gen.__path__` is not a directory (in a frozen
    executable it lives in a zip archive), generated modules are only
    created in memory without writing them to the file system.

    Example:
        GetModule("UIAutomationCore.dll")

    would create modules named

        `comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_L_M_m`
          - typelib wrapper module
          - where L, M, m are numbers of Lcid, Major-ver, minor-ver
        `comtypes.gen.UIAutomationClient`
          - friendly named module

    containing the Python wrapper code for the type library used by
    UIAutomation.  The former module contains all the code, the
    latter is a short stub loading the former.
    """
    if isinstance(tlib, base_text_type):
        tlib_string = tlib
        # if a relative pathname is used, we try to interpret it relative to
        # the directory of the calling module (if not from command line)
        frame = sys._getframe(1)
        _file_ = frame.f_globals.get("__file__", None)  # type: str
        pathname, is_abs = _resolve_filename(tlib_string, _file_ and os.path.dirname(_file_))
        logger.debug("GetModule(%s), resolved: %s", pathname, is_abs)
        tlib = _load_tlib(pathname)  # don't register
        if not is_abs:
            # try to get path after loading, but this only works if already registered
            pathname = tlbparser.get_tlib_filename(tlib)
            if pathname is None:
                logger.info("GetModule(%s): could not resolve to a filename", tlib)
                pathname = tlib_string
        # if above path torture resulted in an absolute path, then the file exists (at this point)!
        assert not(os.path.isabs(pathname)) or os.path.exists(pathname)
    else:
        pathname = None
        tlib = _load_tlib(tlib)
    logger.debug("GetModule(%s)", tlib.GetLibAttr())
    # create and import the real typelib wrapper module
    mod = _create_wrapper_module(tlib, pathname)
    # try to get the friendly-name, if not, returns the real typelib wrapper module
    modulename = codegenerator.name_friendly_module(tlib)
    if modulename is None:
        return mod
    if sys.version_info < (3, 0):
        modulename = modulename.encode("mbcs")
    # create and import the friendly-named module
    return _create_friendly_module(tlib, modulename)


def _load_tlib(obj):
    # type: (Any) -> typeinfo.ITypeLib
    """Load a pointer of ITypeLib on demand."""
    # obj is a filepath or a ProgID
    if isinstance(obj, base_text_type):
        # in any case, attempt to load and if tlib_string is not valid, then raise
        # as "OSError: [WinError -2147312566] Error loading type library/DLL"
        return typeinfo.LoadTypeLibEx(obj)
    # obj is a tlib GUID contain a clsid
    elif isinstance(obj, GUID):
        clsid = str(obj)
        # lookup associated typelib in registry
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\TypeLib" % clsid) as key:
            libid = winreg.EnumValue(key, 0)[1]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\Version" % clsid) as key:
            ver = winreg.EnumValue(key, 0)[1].split(".")
        return typeinfo.LoadRegTypeLib(GUID(libid), int(ver[0]), int(ver[1]), 0)
    # obj is a sequence containing libid
    elif isinstance(obj, (tuple, list)):
        libid, ver = obj[0], obj[1:]
        if not ver:  # case of version numbers are not containing
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"TypeLib\%s" % libid) as key:
                ver = [int(v, base=16) for v in winreg.EnumKey(key, 0).split(".")]
        return typeinfo.LoadRegTypeLib(GUID(libid), *ver)
    # obj is a COMObject implementation
    elif hasattr(obj, "_reg_libid_"):
        return typeinfo.LoadRegTypeLib(GUID(obj._reg_libid_), *obj._reg_version_)
    # obj is a pointer of ITypeLib
    elif isinstance(obj, ctypes.POINTER(typeinfo.ITypeLib)):
        return obj  # type: ignore
    raise TypeError("'%r' is not supported type for loading typelib" % obj)


def _create_module_in_file(modulename, code):
    # type: (str, str) -> types.ModuleType
    """create module in file system, and import it"""
    # `modulename` is 'comtypes.gen.xxx'
    filename = "%s.py" % modulename.split(".")[-1]
    with open(os.path.join(comtypes.client.gen_dir, filename), "w") as ofi:
        print(code, file=ofi)
    # clear the import cache to make sure Python sees newly created modules
    if hasattr(importlib, "invalidate_caches"):
        importlib.invalidate_caches()
    return _my_import(modulename)


def _create_module_in_memory(modulename, code):
    # type: (str, str) -> types.ModuleType
    """create module in memory system, and import it"""
    # `modulename` is 'comtypes.gen.xxx'
    import comtypes.gen as g
    mod = types.ModuleType(modulename)
    abs_gen_path = os.path.abspath(g.__path__[0])  # type: ignore
    mod.__file__ = os.path.join(abs_gen_path, "<memory>")
    exec(code, mod.__dict__)
    sys.modules[modulename] = mod
    setattr(g, modulename.split(".")[-1], mod)
    return mod


def _create_friendly_module(tlib, modulename):
    # type: (typeinfo.ITypeLib, str) -> types.ModuleType
    """helper which creates and imports the friendly-named module."""
    try:
        mod = _my_import(modulename)
    except Exception as details:
        logger.info("Could not import %s: %s", modulename, details)
    else:
        return mod
    # the module is always regenerated if the import fails
    logger.info("# Generating %s", modulename)
    # determine the Python module name
    modname = codegenerator.name_wrapper_module(tlib).split(".")[-1]
    code = "from comtypes.gen import %s\n" % modname
    code += "globals().update(%s.__dict__)\n" % modname
    code += "__name__ = '%s'" % modulename
    if comtypes.client.gen_dir is None:
        return _create_module_in_memory(modulename, code)
    return _create_module_in_file(modulename, code)


def _create_wrapper_module(tlib, pathname):
    # type: (typeinfo.ITypeLib, Optional[str]) -> types.ModuleType
    """helper which creates and imports the real typelib wrapper module."""
    modulename = codegenerator.name_wrapper_module(tlib)
    if modulename in sys.modules:
        return sys.modules[modulename]
    try:
        return _my_import(modulename)
    except Exception as details:
        logger.info("Could not import %s: %s", modulename, details)
    # generate the module since it doesn't exist or is out of date
    logger.info("# Generating %s", modulename)
    p = tlbparser.TypeLibParser(tlib)
    if pathname is None:
        pathname = tlbparser.get_tlib_filename(tlib)
    items = list(p.parse().values())
    codegen = codegenerator.CodeGenerator(_get_known_symbols())
    code = codegen.generate_code(items, filename=pathname)
    for ext_tlib in codegen.externals:  # generates dependency COM-lib modules
        GetModule(ext_tlib)
    if comtypes.client.gen_dir is None:
        return _create_module_in_memory(modulename, code)
    return _create_module_in_file(modulename, code)


def _get_known_symbols():
    # type: () -> Dict[str, str]
    known_symbols = {}  # type: Dict[str, str]
    for mod_name in (
        "comtypes.persist",
        "comtypes.typeinfo",
        "comtypes.automation",
        "comtypes",
        "ctypes.wintypes",
        "ctypes"
    ):
        mod = importlib.import_module(mod_name)
        if hasattr(mod, "__known_symbols__"):
            names = mod.__known_symbols__  # type: List[str]
        else:
            names = list(mod.__dict__)
        for name in names:
            known_symbols[name] = mod.__name__
    return known_symbols

################################################################


if __name__ == "__main__":
    # When started as script, generate typelib wrapper from .tlb file.
    GetModule(sys.argv[1])
