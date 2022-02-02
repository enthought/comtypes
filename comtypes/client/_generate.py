import types
import os
import sys
import comtypes
import comtypes.client
import comtypes.tools.codegenerator
import comtypes.tools.tlbparser
import importlib

import logging
logger = logging.getLogger(__name__)

if sys.version_info >= (3, 0):
    base_text_type = str
else:
    base_text_type = basestring

PATH = os.environ["PATH"].split(os.pathsep)

def _my_import(fullname):
    # helper function to import dotted modules
    import comtypes.gen
    if comtypes.client.gen_dir \
           and comtypes.client.gen_dir not in comtypes.gen.__path__:
        comtypes.gen.__path__.append(comtypes.client.gen_dir)
    return __import__(fullname, globals(), locals(), ['DUMMY'])

def _name_module(tlib):
    # Determine the name of a typelib wrapper module.
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

      (abspath, True) or (relpath, False), 
    
    where relpath is an unresolved path."""
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

    It is possible to delete the whole comtypes\gen directory to
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
    pathname = None
    if isinstance(tlib, base_text_type):
        tlib_string = tlib
        # if a relative pathname is used, we try to interpret it relative to the 
        # directory of the calling module (if not from command line)
        frame = sys._getframe(1)
        _file_ = frame.f_globals.get("__file__", None)
        pathname, path_exists = _resolve_filename(tlib, _file_ and os.path.dirname(_file_))
        logger.debug("GetModule(%s), resolved: %s", pathname, path_exists)
        # in any case, attempt to load and if tlib_string is not valid, then raise
        # as "OSError: [WinError -2147312566] Error loading type library/DLL"
        tlib = comtypes.typeinfo.LoadTypeLibEx(pathname) # don't register
        if not path_exists:
            # try to get path after loading, but this only works if already registered            
            pathname = comtypes.tools.tlbparser.get_tlib_filename(tlib)
            if pathname is None:
                logger.info("GetModule(%s): could not resolve to a filename", tlib)
                pathname = tlib_string
        # if above path torture resulted in an absolute path, then the file exists (at this point)!
        assert not(os.path.isabs(pathname)) or os.path.exists(pathname)
    elif isinstance(tlib, comtypes.GUID):
        # tlib contain a clsid
        clsid = str(tlib)
        
        # lookup associated typelib in registry
        if sys.version_info >= (3, 0):
            import winreg
        else:
            import _winreg as winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\TypeLib" % clsid, 0, winreg.KEY_READ) as key:
            typelib = winreg.EnumValue(key, 0)[1]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\%s\Version" % clsid, 0, winreg.KEY_READ) as key:
            version = winreg.EnumValue(key, 0)[1].split(".")
        
        logger.debug("GetModule(%s)", typelib)
        tlib = comtypes.typeinfo.LoadRegTypeLib(comtypes.GUID(typelib), int(version[0]), int(version[1]), 0)
    elif isinstance(tlib, (tuple, list)):
        # sequence containing libid and version numbers
        logger.debug("GetModule(%s)", (tlib,))
        tlib = comtypes.typeinfo.LoadRegTypeLib(comtypes.GUID(tlib[0]), *tlib[1:])
    elif hasattr(tlib, "_reg_libid_"):
        # a COMObject implementation
        logger.debug("GetModule(%s)", tlib)
        tlib = comtypes.typeinfo.LoadRegTypeLib(comtypes.GUID(tlib._reg_libid_),
                                                *tlib._reg_version_)
    else:
        # an ITypeLib pointer
        logger.debug("GetModule(%s)", tlib.GetLibAttr())

    # create and import the module
    mod = _CreateWrapper(tlib, pathname)
    try:
        modulename = tlib.GetDocumentation(-1)[0]
    except comtypes.COMError:
        return mod
    if modulename is None:
        return mod
    if sys.version_info < (3, 0):
        modulename = modulename.encode("mbcs")

    # create and import the friendly-named module
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
    ofi = open(os.path.join(comtypes.client.gen_dir, modulename + ".py"), "w")
    ofi.write(code)
    ofi.close()
    # clear the import cache to make sure Python sees newly created modules
    if hasattr(importlib, "invalidate_caches"):
        importlib.invalidate_caches()
    return _my_import("comtypes.gen." + modulename)

def _CreateWrapper(tlib, pathname):
    # helper which creates and imports the real typelib wrapper module.
    fullname = _name_module(tlib)
    try:
        return sys.modules[fullname]
    except KeyError:
        pass

    modname = fullname.split(".")[-1]

    try:
        return _my_import(fullname)
    except Exception as details:
        logger.info("Could not import %s: %s", fullname, details)

    # generate the module since it doesn't exist or is out of date
    from comtypes.tools.tlbparser import generate_module
    if comtypes.client.gen_dir is None:
        if sys.version_info >= (3, 0):
            import io
        else:
            import cStringIO as io
        ofi = io.StringIO()
    else:
        ofi = open(os.path.join(comtypes.client.gen_dir, modname + ".py"), "w")
    # XXX use logging!
    logger.info("# Generating comtypes.gen.%s", modname)
    generate_module(tlib, ofi, pathname)

    if comtypes.client.gen_dir is None:
        code = ofi.getvalue()
        mod = types.ModuleType(fullname)
        mod.__file__ = os.path.join(os.path.abspath(comtypes.gen.__path__[0]),
                                    "<memory>")
        exec(code, mod.__dict__)
        sys.modules[fullname] = mod
        setattr(comtypes.gen, modname, mod)
    else:
        ofi.close()
        # clear the import cache to make sure Python sees newly created modules
        if hasattr(importlib, "invalidate_caches"):
            importlib.invalidate_caches()
        mod = _my_import(fullname)
    return mod

################################################################

if __name__ == "__main__":
    # When started as script, generate typelib wrapper from .tlb file.
    GetModule(sys.argv[1])
