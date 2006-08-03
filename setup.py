r"""comtypes - Python COM package, based on the ctypes FFI library.

comtypes allows to define, call, and implement custom COM interfaces
in pure Python.

----------------------------------------------------------------"""
VERSION = "0.3.0"

from distutils.core import setup

##import comtypes

classifiers = [
##    'Development Status :: 3 - Alpha',
    'Development Status :: 4 - Beta',
##    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft :: Windows',
    'Operating System :: Microsoft :: Windows CE',
    'Programming Language :: Python',
    'Topic :: Software Development :: Libraries :: Python Modules',
    ]

setup(name="comtypes",
      description="Pure Python COM package, based on the ctypes package",
      long_description = __doc__,
      author="Thomas Heller",
      author_email="theller@python.net",
      license="MIT License",
      url="http://starship.python.net/crew/theller/comtypes/",

      package_data = {"comtypes.test": ["TestComServer.idl",
                                        "TestComServer.tlb"]},

      classifiers=classifiers,
      
##      version=comtypes.__version__,
      version = VERSION,
      packages=["comtypes",
                "comtypes.client",
                "comtypes.server",
                "comtypes.tools",
                "comtypes.test"])
