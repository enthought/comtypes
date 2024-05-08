import importlib
import os
import sys
import tempfile
import types
import unittest

import comtypes
import comtypes.client
import comtypes.gen


IMGBASE = os.path.splitext(os.path.basename(sys.executable))[0]


class Test(unittest.TestCase):
    """Test the comtypes.client._find_gen_dir() function in several
    simulated environments.
    """

    def setUp(self):
        # save the original comtypes.gen modules and create a
        # substitute with an empty __path__.
        self.orig_comtypesgen = sys.modules["comtypes.gen"]
        del sys.modules["comtypes.gen"]
        del comtypes.gen
        mod = sys.modules["comtypes.gen"] = types.ModuleType("comtypes.gen")
        mod.__path__ = []
        comtypes.gen = mod

    def tearDown(self):
        # Delete py2exe-attributes that we have attached to the sys module
        for name in "frozen frozendllhandle".split():
            try:
                delattr(sys, name)
            except AttributeError:
                pass
        # restore the original comtypes.gen module
        comtypes.gen = self.orig_comtypesgen
        sys.modules["comtypes.gen"] = self.orig_comtypesgen
        importlib.reload(comtypes.gen)

    def test_script(self):
        # %APPDATA%\Python\Python25\comtypes_cache
        ma, mi = sys.version_info[:2]
        cache = rf"$APPDATA\Python\Python{ma:d}{mi:d}\comtypes_cache"
        path = os.path.expandvars(cache)
        gen_dir = comtypes.client._find_gen_dir()
        self.assertEqual(path, gen_dir)

    def test_frozen_dll(self):
        sys.frozen = "dll"
        sys.frozendllhandle = sys.dllhandle
        # %TEMP%\comtypes_cache\<imagebasename>25-25
        # the image is python25.dll
        ma, mi = sys.version_info[:2]
        cache = rf"comtypes_cache\{IMGBASE}{ma:d}{mi:d}-{ma:d}{mi:d}"
        path = os.path.join(tempfile.gettempdir(), cache)
        gen_dir = comtypes.client._find_gen_dir()
        self.assertEqual(path, gen_dir)

    def test_frozen_console_exe(self):
        sys.frozen = "console_exe"
        # %TEMP%\comtypes_cache\<imagebasename>-25
        ma, mi = sys.version_info[:2]
        cache = rf"comtypes_cache\{IMGBASE}-{ma:d}{mi:d}"
        path = os.path.join(tempfile.gettempdir(), cache)
        gen_dir = comtypes.client._find_gen_dir()
        self.assertEqual(path, gen_dir)

    def test_frozen_windows_exe(self):
        sys.frozen = "windows_exe"
        # %TEMP%\comtypes_cache\<imagebasename>-25
        ma, mi = sys.version_info[:2]
        cache = rf"comtypes_cache\{IMGBASE}-{ma:d}{mi:d}"
        path = os.path.join(tempfile.gettempdir(), cache)
        gen_dir = comtypes.client._find_gen_dir()
        self.assertEqual(path, gen_dir)


def main():
    unittest.main()


if __name__ == "__main__":
    main()
