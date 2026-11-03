"""Tests for package installation source verification.

When performing type checking, a dedicated virtual environment is created and
the built package is installed there.
This test suite serves as a assumption check to ensure tests run against
the installed package in site-packages rather than the local source directory
in the workspace.
"""

import site
import unittest as ut
from pathlib import Path

import comtypes


class TestPackaging(ut.TestCase):
    """Verify that comtypes is imported from site-packages."""

    def test_comtypes_is_imported_from_site_packaging(self):
        """Ensure comtypes is imported from site-packages, not local source."""
        init_file = Path(comtypes.__file__).resolve()
        # Verify that comtypes is not loaded from the local repository
        # directory parallel to the tests, which could be included in sys.path.
        self.assertNotEqual(
            init_file, (Path.cwd() / "comtypes" / "__init__.py").resolve()
        )
        # Verify that comtypes is loaded from one of the site-packages
        # directories of the virtual environment.
        self.assertIn(
            init_file,
            [
                (Path(s) / "comtypes" / "__init__.py").resolve()
                for s in site.getsitepackages()
            ],
        )
