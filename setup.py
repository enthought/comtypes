r"""comtypes - pure Python COM package, based on the ctypes FFI library

comtypes offers superior support for custom COM interfaces.

Currently only COM client code is implemented, server support
will follow later.

Limitations:
 - dispinterface support is somewhat weak.

----------------------------------------------------------------"""
VERSION = "0.3.0"

from distutils.core import setup

##import comtypes

classifiers = [
    'Development Status :: 3 - Alpha',
##    'Development Status :: 4 - Beta',
##    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft :: Windows',
    'Programming Language :: Python',
    'Topic :: Software Development :: Libraries :: Python Modules',
    ]

setup(name="comtypes",
      description="pure Python COM package, based on the ctypes FFI library",
      long_description = __doc__,
      author="Thomas Heller",
      author_email="theller@python.net",
      license="MIT License",
      url="http://starship.python.net/crew/theller/comtypes/",

      classifiers=classifiers,
      
##      version=comtypes.__version__,
      version = VERSION,
      packages=["comtypes",
                "comtypes.client",
                "comtypes.server",
                "comtypes.tools",
                "comtypes.test"])
