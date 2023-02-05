from typing import Any, Dict, List, NamedTuple, Optional, Sequence, Tuple, Type
from ctypes import HRESULT, POINTER, c_bool, c_ulong, c_wchar_p
from itertools import permutations
import unittest as ut
from unittest.mock import MagicMock

import comtypes
from comtypes.client import IUnknown
from comtypes._memberspec import _fix_inout_args, _ParamFlagType

WSTRING = c_wchar_p


class Param(NamedTuple):
    argtype: Type
    paramflags: _ParamFlagType


class TestEntry:
    def __init__(self, test_case: ut.TestCase, param_spec: Sequence[Param]):
        self.test_case = test_case
        self.param_spec = param_spec

    def run_test_with_args(self, *args, **kwargs) -> Tuple[Any, MagicMock]:
        """Runs the test with the provided arguments."""
        m = MagicMock()

        argtypes = tuple(x.argtype for x in self.param_spec)
        paramflags = tuple(x.paramflags for x in self.param_spec)
        # fixed_fn = comtypes.instancemethod(
        #     _fix_inout_args(mock_function, argtypes, paramflags), self, None
        # )

        out_params = tuple(
            MagicMock(spec=x.argtype, name=str(x.argtype))
            for x in self.param_spec
            if x.paramflags[0] & 2 == 2
        )
        if len(out_params) == 0:
            out_params = MagicMock(spec=c_ulong, name="HRESULT")
        elif len(out_params) == 1:
            out_params = out_params[0]
        m.return_value = out_params

        fixed_fn = _fix_inout_args(m, argtypes, paramflags)
        # must pass self here because _fix_inout_args expects an instance method
        result = fixed_fn(self, *args, **kwargs)

        return (result, m)

    def run_test(self):
        """Runs the test with automatically generated positional arguments"""
        args = [x.argtype() for x in self.param_spec if x.paramflags[0] & 1 == 1]
        results, mock = self.run_test_with_args(*args)
        mock.assert_called_once_with(self, *args)
        out_params = [x for x in self.param_spec if x.paramflags[0] & 2 == 2]
        if len(out_params) == 0:
            self.test_case.assertIsNone(results)
            return
        if len(out_params) == 1:
            results = [results]
        self.test_case.assertEqual(len(results), len(out_params))
        for result, param in zip(results, out_params):
            if param.paramflags[0] & 3 == 3:
                # inout parameters should be passed back unmodified
                self.test_case.assertIsInstance(result, param.argtype)
            else:
                # out parameters should be generated as MagicMock's
                self.test_case.assertIsInstance(result, MagicMock)


class Test_InOut_args(ut.TestCase):
    # Right now this test fails due to the issue discussed in _memberspec.py
    @ut.expectedFailure
    def test_real_world_examples(self):
        """Test the signatures of several real COM functions"""
        testCases = [
            # IRecordInfo::GetFieldNames
            TestEntry(
                self,
                [
                    Param(POINTER(c_ulong), (3, "pcNames")),
                    Param(POINTER(comtypes.BSTR), (1, "rgBstrNames")),
                ],
            ),
            # ITypeLib::IsName
            TestEntry(
                self,
                [
                    Param(POINTER(c_wchar_p), (3, "name")),
                    Param(c_ulong, (17, "lHashVal", 0)),
                    # the last (out) argument has no name in the header
                    Param(POINTER(c_ulong), (2, None)),
                ],
            ),
            # based on IPortableDeviceContent::CreateObjectWithPropertiesAndData
            # which had a bug in the past
            TestEntry(
                self,
                [
                    Param(POINTER(IUnknown), (1, "pValues")),
                    Param(POINTER(IUnknown), (2, "ppData")),
                    Param(POINTER(c_ulong), (3, "pdwOptimalWriteBufferSize")),
                    Param(POINTER(c_wchar_p), (3, "ppszCookie")),
                ],
            ),
        ]
        for i, entry in enumerate(testCases):
            with self.subTest(f"Example {i}"):
                entry.run_test()

    def test_permutations(self):
        """Test any order of an 'in', an 'out', and two 'inout' arguments of different types"""
        params: List[Param] = [
            Param(c_ulong, (1, "inpar")),
            Param(POINTER(c_wchar_p), (3, "inoutpar1")),
            Param(POINTER(comtypes.IUnknown), (3, "inoutpar2")),
            Param(POINTER(c_bool), (2, "outpar")),
        ]
        for i, permuted_params in enumerate(permutations(params)):
            with self.subTest(f"Permutation {i:#02d}"):
                TestEntry(self, permuted_params).run_test()

    def test_missing_direction(self):
        """Every parameter must have 'in' or 'out' specified"""
        with self.assertRaises(Exception) as cm:
            TestEntry(self, [Param(c_ulong, (0, "missing_inout"))]).run_test()
        self.assertRegex(
            cm.exception.args[0],
            r"^A parameter for .+ has neither 'out' nor 'in' specified$",
        )

    def test_inout_param_name_omitted(self):
        """_fix_inout_args generates a default value for every omitted 'inout' argument."""
        result, mock = TestEntry(
            self,
            [
                Param(
                    POINTER(c_ulong),
                    (3, "param_name"),
                )
            ],
        ).run_test_with_args()
        mock.assert_called_once()
        self.assertEqual(len(mock.call_args[0]), 1)
        self.assertIsInstance(mock.call_args[0][0], TestEntry)
        self.assertEqual(tuple(mock.call_args[1]), ("param_name",))
        generated_arg = mock.call_args[1]["param_name"]
        self.assertIsInstance(generated_arg, c_ulong)
        self.assertEqual(generated_arg.value, c_ulong().value)
        # TODO Not sure what to test 'result' against - right now it is a MagicMock,
        # but I'm not sure it is supposed to be - see my comment in _memberspec.py
        #
        # self.assertIsInstance(result, MagicMock) # works, but seems wrong

    def test_missing_name_omitted(self):
        """The former only works if the argument has a name, so the value can be passed as a keyword argument."""
        with self.assertRaises(Exception) as cm:
            TestEntry(
                self,
                [
                    Param(
                        POINTER(c_ulong),
                        (3, None),
                    )
                ],
            ).run_test_with_args()
        self.assertEqual(
            cm.exception.args[0], "Unnamed inout parameters cannot be omitted"
        )

    def test_inout_params_as_keywords(self):
        """Passing inout parameters as keywords"""
        test_ulong = 5
        test_str = "abc"
        test_p_IUnknown = POINTER(IUnknown)()
        params = [
            Param(POINTER(c_ulong), (3, "long_param")),
            Param(POINTER(c_wchar_p), (3, "str_param")),
            Param(POINTER(IUnknown), (3, "IUnknown_param")),
        ]
        results, mock = TestEntry(self, params).run_test_with_args(
            long_param=test_ulong, str_param=test_str, IUnknown_param=test_p_IUnknown
        )

        self.assertEqual(len(results), 3)
        self.assertEqual(results[0], test_ulong)
        self.assertEqual(results[1], test_str)
        self.assertEqual(results[2], test_p_IUnknown)
        mock.assert_called_once()

        internal_kwargs: Dict[str, Any] = mock.call_args[1]
        self.assertEqual(
            set(internal_kwargs), {"long_param", "str_param", "IUnknown_param"}
        )
        internal_long = internal_kwargs["long_param"]
        internal_str = internal_kwargs["str_param"]
        internal_p_IUnknown = internal_kwargs["IUnknown_param"]
        # Simple types are not passed as pointers
        self.assertIsInstance(internal_long, c_ulong)
        self.assertIsInstance(internal_str, c_wchar_p)
        self.assertIsInstance(internal_p_IUnknown, POINTER(IUnknown))
        self.assertEqual(internal_long.value, test_ulong)
        self.assertEqual(internal_str.value, test_str)
        self.assertEqual(internal_p_IUnknown, test_p_IUnknown)

    def _get_CreateObjectWithPropertiesAndData_spec(self):
        # a memberspec of `PortableDeviceApiLib.IPortableDeviceContent`
        return comtypes.COMMETHOD(
            [],
            HRESULT,
            "CreateObjectWithPropertiesAndData",
            (["in"], POINTER(IUnknown), "pValues"),  # IPortableDeviceValues
            (["out"], POINTER(POINTER(IUnknown)), "ppData"),  # IStream
            (["in", "out"], POINTER(c_ulong), "pdwOptimalWriteBufferSize"),
            (["in", "out"], POINTER(WSTRING), "ppszCookie"),
        )

    def test_CreateObjectWithPropertiesAndData_PositionalsOnly(self):
        spec = self._get_CreateObjectWithPropertiesAndData_spec()
        orig = MagicMock()
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        p_val = MagicMock(
            spec=POINTER(IUnknown), name="POINTER(IPortableDeviceValues)"
        )()

        self_ = MagicMock(name="Self")
        pp_data = POINTER(POINTER(IUnknown))()
        buf_size = 5
        cookie = "abc"
        orig.return_value = (pp_data, ..., ...)
        ret_val = fixed(self_, p_val, buf_size, cookie)
        self.assertEqual(ret_val, [pp_data, buf_size, cookie])
        orig.assert_called_once()
        (orig_0th, orig_1st, orig_2nd, orig_3rd), orig_kw = orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertIsInstance(orig_2nd, c_ulong)
        self.assertEqual(orig_2nd.value, buf_size)
        self.assertIsInstance(orig_3rd, WSTRING)
        self.assertEqual(orig_3rd.value, cookie)
        self.assertEqual(orig_kw, {})

    def test_CreateObjectWithPropertiesAndData_KeywordsOnly(self):
        spec = self._get_CreateObjectWithPropertiesAndData_spec()
        orig = MagicMock()
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        p_val = MagicMock(
            spec=POINTER(IUnknown), name="POINTER(IPortableDeviceValues)"
        )()

        self_ = MagicMock(name="Self")
        pp_data = POINTER(POINTER(IUnknown))()
        buf_size = 5
        cookie = "abc"
        orig.return_value = (pp_data, ..., ...)
        ret_val = fixed(
            self_, pValues=p_val, pdwOptimalWriteBufferSize=buf_size, ppszCookie=cookie
        )
        self.assertEqual(ret_val, [pp_data, buf_size, cookie])
        orig.assert_called_once()
        (orig_0th,), orig_kw = orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(
            set(orig_kw), {"pValues", "pdwOptimalWriteBufferSize", "ppszCookie"}
        )
        self.assertEqual(orig_kw["pValues"], p_val)
        self.assertIsInstance(orig_kw["pdwOptimalWriteBufferSize"], c_ulong)
        self.assertEqual(orig_kw["pdwOptimalWriteBufferSize"].value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    def test_CreateObjectWithPropertiesAndData_MixedArgs_1(self):
        spec = self._get_CreateObjectWithPropertiesAndData_spec()
        orig = MagicMock()
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        p_val = MagicMock(
            spec=POINTER(IUnknown), name="POINTER(IPortableDeviceValues)"
        )()

        self_ = MagicMock(name="Self")
        pp_data = POINTER(POINTER(IUnknown))()
        buf_size = 5
        cookie = "abc"
        orig.return_value = (pp_data, ..., ...)
        ret_val = fixed(
            self_, p_val, ppszCookie=cookie, pdwOptimalWriteBufferSize=buf_size
        )
        self.assertEqual(ret_val, [pp_data, buf_size, cookie])
        orig.assert_called_once()
        (orig_0th, orig_1st), orig_kw = orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertEqual(set(orig_kw), {"pdwOptimalWriteBufferSize", "ppszCookie"})
        self.assertIsInstance(orig_kw["pdwOptimalWriteBufferSize"], c_ulong)
        self.assertEqual(orig_kw["pdwOptimalWriteBufferSize"].value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    def test_CreateObjectWithPropertiesAndData_MixedArgs_2(self):
        spec = self._get_CreateObjectWithPropertiesAndData_spec()
        orig = MagicMock()
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        p_val = MagicMock(
            spec=POINTER(IUnknown), name="POINTER(IPortableDeviceValues)"
        )()

        self_ = MagicMock(name="Self")
        pp_data = POINTER(POINTER(IUnknown))()
        buf_size = 5
        cookie = "abc"
        orig.return_value = (pp_data, ..., ...)
        ret_val = fixed(self_, p_val, buf_size, ppszCookie=cookie)
        self.assertEqual(ret_val, [pp_data, buf_size, cookie])
        orig.assert_called_once()
        (orig_0th, orig_1st, orig_2nd), orig_kw = orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertEqual(set(orig_kw), {"ppszCookie"})
        self.assertIsInstance(orig_2nd, c_ulong)
        self.assertEqual(orig_2nd.value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    # TODO The following lines are not covered by this unit test:
    #
    # else:
    #   v = atyp.from_param(v)
    #   assert not isinstance(v, BYREFTYPE)
    #
    # If you have a natural use case for these lines, please consider adding a test.


if __name__ == "__main__":
    ut.main()
