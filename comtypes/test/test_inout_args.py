from typing import List, NamedTuple
from ctypes import POINTER, c_bool, c_ulong, c_wchar_p
from itertools import permutations
import unittest as ut

import comtypes
from comtypes.client import IUnknown
from comtypes._memberspec import _fix_inout_args, _ParamFlagType


class TestEntry(NamedTuple):
    argtypes: List[type]
    paramflags: List[_ParamFlagType]

    def zipped(self):
        yield from zip(self.argtypes, self.paramflags)

    def run_test(self):
        # TODO future tests could also test keyword or mixed positional/keyword arguments
        def dummy_function(self, *args):
            # return the inout values unmodified, and a dummy output for every pure out value
            result = []
            arg_counter = 0
            for argtype, param in self.zipped():
                if param[0] & 3 == 3:  # inout
                    result.append(args[arg_counter])
                    arg_counter += 1
                elif param[0] & 2 == 2:  # out
                    result.append(argtype())
                elif param[0] & 1 == 1:  # in
                    arg_counter += 1
            if len(result) == 0:
                return None
            elif len(result) == 1:
                return result[0]
            else:
                return tuple(result)

        fixed_fn = comtypes.instancemethod(
            _fix_inout_args(
                dummy_function, tuple(self.argtypes), tuple(self.paramflags)
            ),
            self,
            None,
        )

        dummy_arguments = (
            argtype() for argtype, paramflags in self.zipped() if paramflags[0] & 1 == 1
        )
        # no assertions, just make sure the function call doesn't crash
        fixed_fn(*dummy_arguments)


class Test_InOut_args(ut.TestCase):
    def test_inout_args(self):
        # real world examples
        testCases = [
            TestEntry(
                # IRecordInfo::GetFieldNames
                [POINTER(c_ulong), POINTER(comtypes.BSTR)],
                [(3, "pcNames"), (1, "rgBstrNames")],
            ),
            TestEntry(
                # ITypeLib::IsName
                [POINTER(c_wchar_p), c_ulong, POINTER(c_ulong)],
                # quite interesting: the last (out) argument has no name in the header
                [(3, "name"), (17, "lHashVal", 0), (2, None)],
            ),
            TestEntry(
                # based on IPortableDeviceContent::CreateObjectWithPropertiesAndData
                [
                    POINTER(IUnknown),
                    POINTER(IUnknown),
                    POINTER(c_ulong),
                    POINTER(c_wchar_p),
                ],
                [
                    (1, "pValues"),
                    (2, "ppData"),
                    (3, "pdwOptimalWriteBufferSize"),
                    (3, "ppszCookie"),
                ],
            ),
        ]
        for i, entry in enumerate(testCases):
            with self.subTest(f"Real world example {i}"):
                entry.run_test()

        # fuzzing: any order of 'in', 'out', and two 'inout' arguments
        params: list[tuple[type, int, str]] = [
            (c_ulong, 1, "inpar"),
            (POINTER(c_wchar_p), 3, "inoutpar1"),
            (POINTER(comtypes.IUnknown), 3, "inoutpar"),
            (POINTER(c_bool), 2, "outpar"),
        ]
        for i, permuted_params in enumerate(permutations(params)):
            with self.subTest(f"Permutations example {i}"):
                TestEntry(
                    [p[0] for p in permuted_params], [p[1:3] for p in permuted_params]
                ).run_test()

        # Edge cases
        # 1) missing 'in' or 'out'
        with self.assertRaises(Exception, msg="A parameter for dummyFunction has neither 'out' nor 'in' specified"):
            TestEntry([c_ulong], [(0, 'missing_inout')]).run_test()
        
        # 2) Omitted name for nameless 'inout' parameter
        def dummy_fn_1(first_param):
            return first_param
        fixed_fn = comtypes.instancemethod(
            _fix_inout_args(
                dummy_fn_1, (POINTER(c_ulong),), ((3, None),)
            ),
            self,
            None,
        )
        with self.assertRaises(Exception, msg="Unnamed inout parameters cannot be omitted"):
            # omit an inout parameter whose name is None
            fixed_fn()

        # 3) Omitted name for named 'inout' parameter
        def dummy_fn_2(*_, **kwargs):
            self.assertIn('param_name', kwargs)
            param = kwargs['param_name']
            self.assertIsInstance(param, c_ulong)
            return param
        fixed_fn = comtypes.instancemethod(
            _fix_inout_args(
                dummy_fn_2, (POINTER(c_ulong),), ((3, 'param_name'),)
            ),
            self,
            None,
        )
        fixed_fn()

        # 4) Pass inout parameter as keyword
        def dummy_fn_3(_, keyword_param):
            self.assertIsInstance(keyword_param, c_ulong)
            return keyword_param
        fixed_fn = comtypes.instancemethod(
            _fix_inout_args(
                dummy_fn_3, (POINTER(c_ulong),), ((3, 'keyword_param'),)
            ),
            self,
            None,
        )
        result = fixed_fn(keyword_param=5)
        self.assertEqual(result, 5)


        # 5) Test from_param
        # def dummy_fn_4(_, text):
        #     return text
        # fixed_fn = comtypes.instancemethod(
        #     _fix_inout_args(
        #         dummy_fn_4, (POINTER(IUnknown),), ((3, 'keyword_param'),)
        #     ),
        #     self,
        #     None,
        # )

        # res = fixed_fn("abc")
        # pass

        

if __name__ == "__main__":
    ut.main()
