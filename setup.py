"""comtypes package install script"""
import sys
import os
import subprocess

from distutils.core import Command
from distutils.command.install import install
from setuptools import setup

from distutils.command.build_py import build_py

with open('README.md') as readme_stream:
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
         "repeat tests to search for refcount leaks (requires "
         "'sys.gettotalrefcount')"),
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


try:
    from wheel.bdist_wheel import bdist_wheel
    class bdist_wheel_win(bdist_wheel):
        def get_tag(self):
            win_plats = (  # Copied the list from DistUtils
                "win-amd64",
                "win32",
                "win-arm64",
                "win-arm32",
            )
            tag = super().get_tag()
            return tuple(tag[:-1]) + (".".join(wp.replace("-", "_") for wp in win_plats),)

except ImportError:
    bdist_wheel_win = None


classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft :: Windows',
    'Programming Language :: Python :: 3.7',
    'Programming Language :: Python :: 3.8',
    'Programming Language :: Python :: 3.9',
    'Programming Language :: Python :: 3.10',
    'Programming Language :: Python :: 3.11',
    'Programming Language :: Python :: 3.12',
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
            filename = os.path.join(
                self.install_scripts, "clear_comtypes_cache.py")
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
    long_description=readme,
    author="Thomas Heller",
    author_email="theller@python.net",
    url="https://github.com/enthought/comtypes",
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

    cmdclass={
        'test': test,
        'build_py': build_py,
        'bdist_wheel': bdist_wheel_win,
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

    platforms=["Windows"],
    python_requires=">=3.7",
)

if __name__ == '__main__':
    dist = setup(**setup_params)
    # Exit with a failure code if only running the tests and they failed
    if dist.commands == ['test']:
        command = dist.command_obj['test']
        sys.exit(command.failure)
