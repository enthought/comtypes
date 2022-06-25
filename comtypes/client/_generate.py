from __future__ import print_function
import types
import os
import sys

import comtypes
from comtypes import GUID
import comtypes.client
import comtypes.tools.codegenerator
import comtypes.tools.tlbparser
from comtypes.typeinfo import LoadRegTypeLib, LoadTypeLibEx
import importlib

import logging
logger = logging.getLogger(__name__)

if sys.version_info >= (3, 0):
    base_text_type = str
    import winreg
    import io
else:
    base_text_type = basestring
    import _winreg as winreg
    import cStringIO as io


PATH = os.environ["PATH"].split(os.pathsep)


def _my_import(fullname):
    """helper function to import dotted modules"""
    import comtypes.gen
    if comtypes.client.gen_dir \
           and comtypes.client.gen_dir not in comtypes.gen.__path__:
        comtypes.gen.__path__.append(comtypes.client.gen_dir)
    return importlib.import_module(fullname)


def _name_module(tlib):
    """Determine the name of a typelib wrapper module"""
    libattr = tlib.GetLibAttr()
    modname = "_%s_%s_%s_%s" % \
              (str(libattr.guid)[1:-1].replace("-", "_"),
               libattr.lcid,
               libattr.wMajorVerNum,
               libattr.wMinorVerNum)
    return "comtypes.gen." + modname


def _resolve_filename(tlib_string, dirpath):
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
    """Create a module wrapping a COM typelibrary on demand.

    'tlib' must be an ITypeLib COM pointer instance, the pathname of a
    type library, a COM CLSID GUID, or a tuple/list specifying the
    arguments to a comtypes.typeinfo.LoadRegTypeLib call:

      (libid, wMajorVerNum, wMinorVerNum, lcid=0)

    Or it can be an object with _reg_libid_ and _reg_version_
    attributes.

    A relative pathname is interpreted as relative to the callers
    __file__, if this exists.

    This function determines the module name from the typelib
    attributes, then tries to import it.  If that fails because the
    module doesn't exist, the module is generated into the
    comtypes.gen package.

    It is possible to delete the whole `comtypes/gen` directory to
    remove all generated modules, the directory and the __init__.py
    file in it will be recreated when needed.

    If comtypes.gen __path__ is not a directory (in a frozen
    executable it lives in a zip archive), generated modules are only
    created in memory without writing them to the file system.

    Example:

        GetModule("shdocvw.dll")

    would create modules named

       comtypes.gen._EAB22AC0_30C1_11CF_A7EB_0000C05BAE0B_0_1_1
       comtypes.gen.SHDocVw

    containing the Python wrapper code for the type library used by
    Internet Explorer.  The former module contains all the code, the
    latter is a short stub loading the former.
    """
    if isinstance(tlib, base_text_type):
        tlib_string = tlib
        # if a relative pathname is used, we try to interpret it relative to the 
        # directory of the calling module (if not from command line)
        frame = sys._getframe(1)
        _file_ = frame.f_globals.get("__file__", None)
        pathname, is_abs = _resolve_filename(tlib_string, _file_ and os.path.dirname(_file_))
        logger.debug("GetModule(%s), resolved: %s", pathname, is_abs)
        tlib = _load_tlib(pathname)  # don't register
        if not is_abs:
            # try to get path after loading, but this only works if already registered            
            pathname = comtypes.tools.tlbparser.get_tlib_filename(tlib)
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
    try:
        modulename = tlib.GetDocumentation(-1)[0]
    except comtypes.COMError:
        return mod
    if modulename is None:
        return mod
    if sys.version_info < (3, 0):
        modulename = modulename.encode("mbcs")
    # create and import the friendly-named module
    return _create_friendly_module(tlib, modulename)


def _load_tlib(obj):
    """Load a pointer of ITypeLib on demand."""
    # obj is a filepath or a ProgID
    if isinstance(obj, base_text_type):
        # in any case, attempt to load and if tlib_string is not valid, then raise
        # as "OSError: [WinError -2147312566] Error loading type library/DLL"
        return LoadTypeLibEx(obj)
    # obj is a tlib GUID contain a clsid
    elif isinstance(obj, GUID):
        clsid = str(obj)
        # lookup associated typelib in registry
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\TypeLib" % clsid) as key:
            libid = winreg.EnumValue(key, 0)[1]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\Version" % clsid) as key:
            version = winreg.EnumValue(key, 0)[1].split(".")
        return LoadRegTypeLib(GUID(libid), int(version[0]), int(version[1]), 0)
    # obj is a sequence containing libid
    elif isinstance(obj, (tuple, list)):
        libid, version = obj[0], obj[1:]
        if not version:  # case of version numbers are not containing
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"TypeLib\%s" % libid) as key:
                version = [int(v, base=16) for v in winreg.EnumKey(key, 0).split(".")]
        return LoadRegTypeLib(GUID(libid), *version)
    # obj is a COMObject implementation
    elif hasattr(obj, "_reg_libid_"):
        return LoadRegTypeLib(GUID(obj._reg_libid_), *obj._reg_version_)
    # perhaps obj is a pointer of ITypeLib
    return obj


def _invalidate_import_caches():
    """clear the import cache to make sure Python sees newly created modules"""
    if hasattr(importlib, "invalidate_caches"):
        importlib.invalidate_caches()


def _create_friendly_module(tlib, modulename):
    """helper which creates and imports the friendly-named module."""
    try:
        mod = _my_import("comtypes.gen." + modulename)
    except Exception as details:
        logger.info("Could not import comtypes.gen.%s: %s", modulename, details)
    else:
        return mod
    # the module is always regenerated if the import fails
    logger.info("# Generating comtypes.gen.%s", modulename)
    # determine the Python module name
    fullname = _name_module(tlib)
    modname = fullname.split(".")[-1]
    code = "from comtypes.gen import %s\nglobals().update(%s.__dict__)\n" % (modname, modname)
    code += "__name__ = 'comtypes.gen.%s'" % modulename
    if comtypes.client.gen_dir is None:
        mod = types.ModuleType("comtypes.gen." + modulename)
        mod.__file__ = os.path.join(os.path.abspath(comtypes.gen.__path__[0]),
                                    "<memory>")
        exec(code, mod.__dict__)
        sys.modules["comtypes.gen." + modulename] = mod
        setattr(comtypes.gen, modulename, mod)
        return mod
    # create in file system, and import it
    with open(os.path.join(comtypes.client.gen_dir, modulename + ".py"), "w") as ofi:
        print(code, file=ofi)
    _invalidate_import_caches()
    return _my_import("comtypes.gen." + modulename)


def _create_wrapper_module(tlib, pathname):
    """helper which creates and imports the real typelib wrapper module."""
    fullname = _name_module(tlib)
    if fullname in sys.modules:
        return sys.modules[fullname]

    modname = fullname.split(".")[-1]

    try:
        return _my_import(fullname)
    except Exception as details:
        logger.info("Could not import %s: %s", fullname, details)

    # generate the module since it doesn't exist or is out of date
    stream = io.StringIO()
    logger.info("# Generating comtypes.gen.%s", modname)
    comtypes.tools.tlbparser.generate_module(tlib, stream, pathname)
    if comtypes.client.gen_dir is None:
        mod = types.ModuleType(fullname)
        mod.__file__ = os.path.join(os.path.abspath(comtypes.gen.__path__[0]),
                                    "<memory>")
        exec(stream.getvalue(), mod.__dict__)
        sys.modules[fullname] = mod
        setattr(comtypes.gen, modname, mod)
    else:
        with open(os.path.join(comtypes.client.gen_dir, modname + ".py"), "w") as ofi:
            print(stream.getvalue(), file=ofi)
        _invalidate_import_caches()
        mod = _my_import(fullname)
    return mod

################################################################

if __name__ == "__main__":
    # When started as script, generate typelib wrapper from .tlb file.
    GetModule(sys.argv[1])
