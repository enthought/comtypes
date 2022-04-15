"""comtypes package install script"""
import sys
import os
import ctypes
import subprocess

from distutils.core import Command
from distutils.command.install import install
from setuptools import setup

from distutils.command.build_py import build_py

with open('README') as readme_stream:
    readme = readme_stream.read()


classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft :: Windows',
    'Programming Language :: Python',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3',
    'Topic :: Software Development :: Libraries :: Python Modules',
    ]

def read_version():
    # Determine the version number by reading it from the file
    # 'comtypes\__init__.py'.  We cannot import this file (with py3,
    # at least) because it is in py2.x syntax.
    for line in open("comtypes/__init__.py"):
        if line.startswith("__version__ = "):
            var, value = line.split('=')
            return value.strip().strip('"').strip("'")
    raise NotImplementedError("__version__ is not found in __init__.py")


class post_install(install):

    # both this static variable and method initialize_options() help to avoid
    # weird setuptools error with "pip install comtypes", details are here:
    # https://github.com/enthought/comtypes/issues/155
    # the working solution was found here:
    # https://github.com/pypa/setuptools/blob/3b90be7bb6323eb44d0f28864509c1d47aa098de/setuptools/command/install.py
    user_options = install.user_options + [
        ('old-and-unmanageable', None, "Try not to use this!"),
        ('single-version-externally-managed', None,
         "used by system package builders to create 'flat' eggs"),
    ]

    def initialize_options(self):
        install.initialize_options(self)
        self.old_and_unmanageable = None
        self.single_version_externally_managed = None

    def run(self):
        install.run(self)
        # Custom script we run at the end of installing
        if not self.dry_run and not self.root:
            filename = os.path.join(self.install_scripts, "clear_comtypes_cache.py")
            if not os.path.isfile(filename):
                raise RuntimeError("Can't find '%s'" % (filename,))
            print("Executing post install script...")
            print('"' + sys.executable + '" "' + filename + '" -y')
            try:
                subprocess.check_call([sys.executable, filename, '-y'])
            except subprocess.CalledProcessError:
                print("Failed to run post install script!")


setup_params = dict(
    name="comtypes",
    description="Pure Python COM package",
    long_description = readme,
    author="Thomas Heller",
    author_email="theller@python.net",
    url="https://github.com/enthought/comtypes",
    download_url="https://github.com/enthought/comtypes/releases",
    license="MIT License",
    classifiers=classifiers,

    scripts=["clear_comtypes_cache.py"],

    cmdclass={
        'build_py': build_py,
        'install': post_install,
    },

    version=read_version(),
    packages=[
        "comtypes",
        "comtypes.client",
        "comtypes.server",
        "comtypes.tools",
        "comtypes.test",
    ],
)

if __name__ == '__main__':
    dist = setup(**setup_params)
