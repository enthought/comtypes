"comtypes package install script"

import sys
import os
import ctypes

from distutils.core import Command
from setuptools import setup

try:
    from distutils.command.build_py import build_py_2to3 as build_py
except ImportError:
    from distutils.command.build_py import build_py

with open('README') as readme_stream:
    readme = readme_stream.read()

class test(Command):
    # Original version of this class posted
    # by Berthold Hoellmann to distutils-sig@python.org
    description = "run tests"

    user_options = [
        ('tests=', 't',
         "comma-separated list of packages that contain test modules"),
        ('use-resources=', 'u',
         "resources to use - resource names are defined by tests"),
        ('refcounts', 'r',
         "repeat tests to search for refcount leaks (requires 'sys.gettotalrefcount')"),
        ]

    boolean_options = ["refcounts"]

    def initialize_options(self):
        self.use_resources = ""
        self.refcounts = False
        self.tests = "comtypes.test"
        self.failure = False

    def finalize_options(self):
        if self.refcounts and not hasattr(sys, "gettotalrefcount"):
            raise Exception("refcount option requires Python debug build")
        self.tests = self.tests.split(",")
        self.use_resources = self.use_resources.split(",")

    def run(self):
        build = self.reinitialize_command('build')
        build.run()
        if build.build_lib is not None:
            sys.path.insert(0, build.build_lib)

        # Register our ATL COM tester dll
        import comtypes.test
        script_path = os.path.dirname(__file__)
        source_dir = os.path.abspath(os.path.join(script_path, "source"))
        comtypes.test.register_server(source_dir)

        comtypes.test.use_resources.extend(self.use_resources)
        for name in self.tests:
            package = __import__(name, globals(), locals(), ['*'])
            sys.stdout.write("Testing package %s %s\n"
                             % (name, (sys.version, sys.platform, os.name)))
            package_failure = comtypes.test.run_tests(package,
                                                      "test_*.py",
                                                      self.verbose,
                                                      self.refcounts)
            self.failure = self.failure or package_failure

classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft :: Windows',
    'Operating System :: Microsoft :: Windows :: Windows CE',
    'Programming Language :: Python',
    'Programming Language :: Python :: 2.6',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3',
    'Topic :: Software Development :: Libraries :: Python Modules',
    ]

def read_version():
    # Determine the version number by reading it from the file
    # 'comtypes\__init__.py'.  We cannot import this file (with py3,
    # at least) because it is in py2.x syntax.
    ns = {}
    for line in open("comtypes/__init__.py"):
        if line.startswith("__version__ = "):
            exec(line, ns)
            break
    return ns["__version__"]

if sys.version_info >= (3, 0):
    # install_script does not work in Python 3 (python bug)
    # Another distutils bug: it doesn't accept an empty options dict
    options = {"foo": {}}
##    options = {}
else:
    options={"bdist_wininst": {"install_script": "clear_comtypes_cache.py"}}

setup_params = dict(
    name="comtypes",
    description="Pure Python COM package",
    long_description = readme,
    author="Thomas Heller",
    author_email="theller@python.net",
    url="http://starship.python.net/crew/theller/comtypes",
    download_url="https://github.com/enthought/comtypes/releases",
    license="MIT License",
    package_data={
        "comtypes.test": [
            "TestComServer.idl",
            "TestComServer.tlb",
            "TestDispServer.idl",
            "TestDispServer.tlb",
            "mytypelib.idl",
            "mylib.idl",
            "mylib.tlb"
            "urlhist.tlb",
            "test_jscript.js",
        ]},
    classifiers=classifiers,

    scripts=["clear_comtypes_cache.py"],
    options=options,

    cmdclass={
        'test': test,
        'build_py': build_py,
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
    # Exit with a failure code if only running the tests and they failed
    if dist.commands == ['test']:
        command = dist.command_obj['test']
        sys.exit(command.failure)
