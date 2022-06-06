"""This test covers 'pip install' issue #155"""
import os
import sys
import shutil
import subprocess
import unittest

def read_version():
    # Determine the version number by reading it from the file
    # 'comtypes\__init__.py'.  We cannot import this file (with py3,
    # at least) because it is in py2.x syntax.
    with open("comtypes/__init__.py") as ofi:
        for line in ofi:
            if line.startswith("__version__ = "):
                var, value = line.split('=')
                return value.strip().strip('"').strip("'")
    raise NotImplementedError("__version__ is not found in __init__.py")


class TestPipInstall(unittest.TestCase):

    def setUp(self):
        """prepare the same package that is usually uploaded to PyPI"""
        subprocess.check_call([sys.executable, 'setup.py', 'sdist', '--format=zip'])

        filename_for_upload = 'comtypes-%s.zip' % read_version()
        self.target_package = os.path.join(os.getcwd(), 'dist', filename_for_upload)
        self.pip_exe = os.path.join(os.path.dirname(sys.executable), 'Scripts', 'pip.exe')

    def test_pip_install(self):
        """Test that "pip install comtypes-x.y.z.zip" works"""
        subprocess.check_call([self.pip_exe, 'install', self.target_package])

    def test_no_cache_dir_custom_location(self):
        """Test that 'pip install comtypes-x.y.z.zip --no-cache-dir --target="...\custom location"' works"""
        custom_dir = os.path.join(os.getcwd(), 'custom location')
        if os.path.exists(custom_dir):
            shutil.rmtree(custom_dir)
        os.makedirs(custom_dir)

        # this test catches issue #158
        subprocess.check_call('{0} install {1} --no-cache-dir --target="{2}"' \
            ''.format(self.pip_exe, self.target_package, custom_dir))


if __name__ == '__main__':
    unittest.main()
