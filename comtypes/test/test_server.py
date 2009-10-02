from comtypes.client import GetModule
from comtypes.hresult import S_OK
import os, sys

def is_newer(src, dst):
    "Checks if file 'src' is newer than file 'dst'"
    if not os.path.isfile(dst):
        return True
    src_mtime = os.stat(src).st_mtime
    dst_mtime = os.stat(dst).st_mtime
    return dst_mtime < src_mtime

curdir = os.path.dirname(__file__)

# Compile idl file into type lib, if needed.  This requires
# MSVC 7.1
if is_newer("%s\mytypelib.idl" % curdir,
            "%s\mytypelib.tlb" % curdir):
    print "Compiling mytypelib.idl into typelib..."
    os.system(r'call "%%VS71COMNTOOLS%%vsvars32.bat" && '
              r'midl /nologo %s\mytypelib.idl /tlb %s\mytypelib.tlb' % (curdir, curdir))

# Register the typelib
from comtypes.typeinfo import LoadTypeLib
# mytypelib.idl
LoadTypeLib(os.path.join(os.path.dirname(__file__), "mytypelib.tlb"))

# Generate the comtypes typelib wrapper
GetModule("mytypelib.tlb")

from comtypes.gen import MyTypeLib

################################################################
# We implement a COM object in this module, but it isn't registered
# with the COM runtime or in the the Registry.  The unittest below
# instantiates the object in an unusual way - however, the method
# calls use the normal COM machinery.

class MyServer(MyTypeLib.MyComServer):

    @property
    def Id(self):
        return id(self)

    def Test(self, value):
        return value * 3
    
    name = "foo"

    def eval(self, what):
        return eval(what)

    def EXEC(self, what):
        exec(what)

    def TestPairArray(self, pval):
        # pval is a POINTER to a SAFEARRAY.
        # pval[0] is then a sequence of the elements

        # we can return a sequence of Pair objects...
        return [MyTypeLib.Pair(3.14, 2.78)]

    def TestPairArray2(self, pval):
        # pval is a POINTER to a SAFEARRAY.
        # pval[0] is then a sequence of the elements

        # we can also return a sequence of number tuples,
        # which will be packed into Pairs.
        return [(42 * len(pval[0]), 43 * len(pval[0]))]

##    HRESULT MultiOutArgs([in, out] int *pa,
##                         [in, out] int *pb,
##                         [in, out] int *pc);
    def MultiInOutArgs(self, pa, pb, pc):
        return 1, 2, 3

    # The [out, retval] parameter is NOT passed!
    def MultiOutArgs2(self, pa, pb):
        return 1, 2, 3

import unittest
from comtypes import IUnknown
from comtypes.client import wrap


class Test(unittest.TestCase):

    def setUp(self):
        obj = MyServer()
        self.p = wrap(obj.QueryInterface(IUnknown))

    def test_basics(self):
        p = self.p
        self.assertEqual(p.Name, "foo")
        self.assertEqual(p.name, "foo")
        p.Name = "bar"
        self.assertEqual(p.Name, "bar")
        self.assertEqual(p.name, "bar")
        self.assertEqual(p.eval("sys.version"), sys.version)
        self.assertEqual(p.Test(42), 42*3)

    def test_Records(self):
        p = self.p
        pairs = p.TestPairArray([MyTypeLib.Pair(1, 2), MyTypeLib.Pair(4, 5)])
        self.assertEqual(pairs[0].a, 3.14)
        self.assertEqual(pairs[0].b, 2.78)

        pairs = p.TestPairArray2([MyTypeLib.Pair(1, 2), MyTypeLib.Pair(4, 5)])
        self.assertEqual(pairs[0].a, 42*2)
        self.assertEqual(pairs[0].b, 43*2)

    def test_inout_args(self):
        p = self.p
        # [in, out] args are optional
        self.assertEqual(p.MultiInOutArgs(), (1, 2, 3))
        self.assertEqual(p.MultiInOutArgs(-1, -1, -1), (1, 2, 3))

    def test_inout_outretval_args(self):
        # This test fails!
        p = self.p
        # BUG: p.MultiOutArgs2() does NOT return a 3-tuple, the third item is lost
        # somewhere
        self.assertEqual(p.MultiOutArgs2(), (1, 2, 3))

if __name__ == "__main__":
    unittest.main()
