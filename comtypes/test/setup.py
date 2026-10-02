# all the unittests can be converted to exe-files.
import glob
from distutils.core import setup


setup(name="test_*", console=glob.glob("test_*.py"))
