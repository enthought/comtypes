"""
Test for the ``comtypes.clear_cache`` module.
"""
import io
import runpy
from unittest.mock import patch, call
from unittest import TestCase

from comtypes.client import _find_gen_dir


class ClearCacheTestCase(TestCase):
    # we patch sys.stdout so unittest doesn't show the print statements
    @patch("sys.stdout", new=io.StringIO())
    @patch("shutil.rmtree")
    def test_clear_cache(self, mock_rmtree):
        with patch("sys.argv", ["clear_cache.py", "-y"]):
            runpy.run_module("comtypes.clear_cache", {}, "__main__")

        # because we don't actually delete anything, _find_gen_dir() will
        # give the same answer every time we call it
        assert mock_rmtree.call_args_list == [call(_find_gen_dir()) for i in range(2)]
