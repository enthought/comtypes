"""comtypes package install script"""

import sys
import os
import subprocess

from setuptools import Command, setup
from setuptools.command.install import install


class test(Command):
    # Original version of this class posted
    # by Berthold Hoellmann to distutils-sig@python.org
    description = "run tests"

    user_options = [
        ('tests=', 't', "comma-separated list of packages that contain test modules"),
        (
            'use-resources=',
            'u',
            "resources to use - resource names are defined by tests",
        ),
        (
            'refcounts',
            'r',
            "repeat tests to search for refcount leaks (requires "
            "'sys.gettotalrefcount')",
        ),
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
            sys.stdout.write(
                "Testing package %s %s\n" % (name, (sys.version, sys.platform, os.name))
            )
            package_failure = comtypes.test.run_tests(
                package, "test_*.py", self.verbose, self.refcounts
            )
            self.failure = self.failure or package_failure


class post_install(install):
    # both this static variable and method initialize_options() help to avoid
    # weird setuptools error with "pip install comtypes", details are here:
    # https://github.com/enthought/comtypes/issues/155
    # the working solution was found here:
    # https://github.com/pypa/setuptools/blob/3b90be7bb6323eb44d0f28864509c1d47aa098de/setuptools/command/install.py
    user_options = install.user_options + [
        ('old-and-unmanageable', None, "Try not to use this!"),
        (
            'single-version-externally-managed',
            None,
            "used by system package builders to create 'flat' eggs",
        ),
    ]

    def initialize_options(self):
        install.initialize_options(self)
        self.old_and_unmanageable = None
        self.single_version_externally_managed = None

    def run(self):
        install.run(self)
        # Custom script we run at the end of installing
        if not self.dry_run and not self.root:
            print("Executing post install script...")
            print(f'"{sys.executable}" -m comtypes.clear_cache -y')
            try:
                subprocess.check_call([
                    sys.executable,
                    "-m",
                    "comtypes.clear_cache",
                    '-y',
                ])
            except subprocess.CalledProcessError:
                print("Failed to run post install script!")


if __name__ == '__main__':
    dist = setup(
        cmdclass={
            'test': test,
            'install': post_install,
        },
    )
    # Exit with a failure code if only running the tests and they failed
    if dist.commands == ['test']:
        command = dist.command_obj['test']
        sys.exit(command.failure)
