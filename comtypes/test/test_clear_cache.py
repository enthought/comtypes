"""
Tests for the ``comtypes.clear_cache`` module.

This module provides tests for the ``clear_cache`` script. Because we don't
want a test invocation of the script to actually delete the real cache of the
calling Python (which would probably break many of the following tests), we use
pyfakefs which creates an in memory mock file system which mirrors the relevant
folders but does not propagate changes to the real version.

Because there are various locations that the cache folder can be in, we include
all the folders from sys.path in the mock fs, as the comtypes.gen module has to
be in one of them.
"""
import io
import os
import runpy
import sys
from unittest.mock import patch

from pyfakefs.fake_filesystem_unittest import TestCase

# importing this will create a real gen cache dir which is necessary as
# comtypes relies on importing the module (and can't import from the fake fs)
import comtypes.client


class ClearCacheTestCase(TestCase):
    def setUp(self) -> None:
        self.setUpPyfakefs()

    # we patch sys.stdout so unittest doesn't show the print statements
    @patch("sys.stdout", new=io.StringIO())
    def test_clear_cache(self):

        for site_path in sys.path:
            try:
                self.fs.add_real_directory(site_path, read_only=False)
            except (FileExistsError, FileNotFoundError):
                pass

        # ask comtypes where the cache dir is so we can check it is gone
        cache_dir = comtypes.client._find_gen_dir()
        assert os.path.exists(cache_dir)

        with patch("sys.argv", ["clear_cache.py", "-y"]):
            runpy.run_module("comtypes.clear_cache", {}, "__main__")

        assert not os.path.exists(cache_dir)
