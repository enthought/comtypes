from typing import Any, Callable, List, NamedTuple, Sequence, Tuple, Type
from ctypes import HRESULT, POINTER, c_bool, c_ulong, c_wchar_p
from itertools import permutations
import unittest as ut
from unittest.mock import MagicMock

import comtypes
from comtypes.client import IUnknown
from comtypes._memberspec import _fix_inout_args, _ParamFlagType, _ArgSpecElmType

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
        params = [
            (["in"], c_ulong, "inpar"),
            (["in", "out"], POINTER(c_wchar_p), "inoutpar1"),
            (["in", "out"], POINTER(comtypes.IUnknown), "inoutpar2"),
            (["out"], POINTER(c_bool), "outpar"),
        ]
        orig = MagicMock()
        for i, permuted_params in enumerate(permutations(params)):
            spec = comtypes.COMMETHOD(
                [], HRESULT, "CreateObjectWithPropertiesAndData", *permuted_params
            )
            with self.subTest(f"Permutation {i:#02d}"):
                fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
                self_ = MagicMock(name="Self")
                generated_in_params = tuple(
                    param[1]() for param in permuted_params if "in" in param[0]
                )
                mock_outpar = MagicMock(name="outpar")
                mock_return_values = []
                for param in permuted_params:
                    if "out" in param[0]:
                        mock_return_values.append(
                            ... if "in" in param[0] else mock_outpar
                        )
                orig.return_value = tuple(mock_return_values)
                ret_vals = fixed(self_, *generated_in_params)

                # These need to be matched
                call_arguments_iter = iter(orig.call_args[0][1:])
                gen_in_vals_iter = iter(generated_in_params)
                ret_vals_iter = iter(ret_vals)
                for param in permuted_params:
                    if param[0] == ["in", "out"]:
                        # check the input into 'fixed' against the input into 'orig' and the output of 'fixed'
                        in_arg = next(call_arguments_iter)
                        self.assertEqual(in_arg, next(ret_vals_iter))
                        self.assertEqual(in_arg, next(gen_in_vals_iter))
                    elif param[0] == ["out"]:
                        # check the output of 'fixed' against the pre-defined mock
                        self.assertIs(next(ret_vals_iter), mock_outpar)
                    else:  # ["in"]
                        # check the input into 'fixed' against the input into 'orig'
                        self.assertEqual(
                            next(call_arguments_iter), next(gen_in_vals_iter)
                        )

    # TODO The following lines are not covered by this unit test:
    #
    # else:
    #   v = atyp.from_param(v)
    #   assert not isinstance(v, BYREFTYPE)
    #
    # If you have a natural use case for these lines, please consider adding a test.


class Test_IPortableDeviceContent_CreateObjectWithPropertiesAndData(ut.TestCase):
    def setUp(self):
        # a memberspec of `PortableDeviceApiLib.IPortableDeviceContent`
        spec = comtypes.COMMETHOD(
            [],
            HRESULT,
            "CreateObjectWithPropertiesAndData",
            (["in"], POINTER(IUnknown), "pValues"),  # IPortableDeviceValues
            (["out"], POINTER(POINTER(IUnknown)), "ppData"),  # IStream
            (["in", "out"], POINTER(c_ulong), "pdwOptimalWriteBufferSize"),
            (["in", "out"], POINTER(WSTRING), "ppszCookie"),
        )
        self.orig = MagicMock()
        self.fixed = _fix_inout_args(self.orig, spec.argtypes, spec.paramflags)
        # out and inout argument are dereferenced once before being returned
        self.pp_data = POINTER(IUnknown)()
        self.orig.return_value = (self.pp_data, ..., ...)

    def test_positionals_only(self):
        self_ = MagicMock(name="Self")
        p_val = MagicMock(spec=POINTER(IUnknown))()
        buf_size = 5
        cookie = "abc"
        ret_val = self.fixed(self_, p_val, buf_size, cookie)

        self.assertEqual(ret_val, [self.pp_data, buf_size, cookie])
        self.orig.assert_called_once()
        (orig_0th, orig_1st, orig_2nd, orig_3rd), orig_kw = self.orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertIsInstance(orig_2nd, c_ulong)
        self.assertEqual(orig_2nd.value, buf_size)
        self.assertIsInstance(orig_3rd, WSTRING)
        self.assertEqual(orig_3rd.value, cookie)
        self.assertEqual(orig_kw, {})

    def test_keywords_only(self):
        self_ = MagicMock(name="Self")
        p_val = MagicMock(spec=POINTER(IUnknown))()
        buf_size = 4
        cookie = "efg"
        ret_val = self.fixed(
            self_, pValues=p_val, pdwOptimalWriteBufferSize=buf_size, ppszCookie=cookie
        )

        self.assertEqual(ret_val, [self.pp_data, buf_size, cookie])
        self.orig.assert_called_once()
        (orig_0th,), orig_kw = self.orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(
            set(orig_kw), {"pValues", "pdwOptimalWriteBufferSize", "ppszCookie"}
        )
        self.assertEqual(orig_kw["pValues"], p_val)
        self.assertIsInstance(orig_kw["pdwOptimalWriteBufferSize"], c_ulong)
        self.assertEqual(orig_kw["pdwOptimalWriteBufferSize"].value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    def test_mixed_args_1(self):
        self_ = MagicMock(name="Self")
        p_val = MagicMock(spec=POINTER(IUnknown))()
        buf_size = 3
        cookie = "h"
        ret_val = self.fixed(
            self_, p_val, ppszCookie=cookie, pdwOptimalWriteBufferSize=buf_size
        )

        self.assertEqual(ret_val, [self.pp_data, buf_size, cookie])
        self.orig.assert_called_once()
        (orig_0th, orig_1st), orig_kw = self.orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertEqual(set(orig_kw), {"pdwOptimalWriteBufferSize", "ppszCookie"})
        self.assertIsInstance(orig_kw["pdwOptimalWriteBufferSize"], c_ulong)
        self.assertEqual(orig_kw["pdwOptimalWriteBufferSize"].value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    def test_mixed_args_2(self):
        self_ = MagicMock(name="Self")
        p_val = MagicMock(spec=POINTER(IUnknown))()
        buf_size = 2
        cookie = "ij"
        ret_val = self.fixed(self_, p_val, buf_size, ppszCookie=cookie)

        self.assertEqual(ret_val, [self.pp_data, buf_size, cookie])
        self.orig.assert_called_once()
        (orig_0th, orig_1st, orig_2nd), orig_kw = self.orig.call_args
        self.assertIs(orig_0th, self_)
        self.assertEqual(orig_1st, p_val)
        self.assertEqual(set(orig_kw), {"ppszCookie"})
        self.assertIsInstance(orig_2nd, c_ulong)
        self.assertEqual(orig_2nd.value, buf_size)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, cookie)

    def test_omitted_arguments_autogen(self):
        self_ = MagicMock(name="Self")
        p_val = MagicMock(spec=POINTER(IUnknown))()
        ret_val = self.fixed(self_, pValues=p_val)

        self.orig.assert_called_once()
        (orig_0th,), orig_kw = self.orig.call_args
        self.assertEqual(
            set(orig_kw), {"pValues", "pdwOptimalWriteBufferSize", "ppszCookie"}
        )
        self.assertIs(orig_0th, self_)
        self.assertIs(orig_kw["pValues"], p_val)
        self.assertIsInstance(orig_kw["pdwOptimalWriteBufferSize"], c_ulong)
        self.assertEqual(orig_kw["pdwOptimalWriteBufferSize"].value, c_ulong().value)
        self.assertIsInstance(orig_kw["ppszCookie"], WSTRING)
        self.assertEqual(orig_kw["ppszCookie"].value, WSTRING().value)

        self.assertEqual(
            ret_val,
            [
                self.pp_data,
                orig_kw["pdwOptimalWriteBufferSize"].value,
                orig_kw["ppszCookie"].value,
            ],
        )


class PermutedArgspecTestingParams(NamedTuple):
    argspec: Tuple[_ArgSpecElmType, _ArgSpecElmType, _ArgSpecElmType, _ArgSpecElmType]
    args: Tuple[Any, Any, Any]
    orig_ret_val: Tuple[Any, Any, Any]
    fixed_ret_val: List[Any]
    call_args_validators: List[Tuple[Type[Any], Callable[[Any], Any], Any]]


class Test_ArgspecPermutations(ut.TestCase):
    def test_permutations(self):
        for testing_params in self._get_params():
            with self.subTest(testing_params):
                self_ = MagicMock(name="Self")
                orig = MagicMock(return_value=testing_params.orig_ret_val)
                spec = comtypes.COMMETHOD([], HRESULT, "foo", *testing_params.argspec)
                fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
                ret_val = fixed(self_, *testing_params.args)

                self.assertEqual(ret_val, testing_params.fixed_ret_val)
                orig.assert_called_once()
                (orig_0th, *orig_call_args), orig_kw = orig.call_args
                self.assertEqual(orig_kw, {})
                self.assertIs(orig_0th, self_)
                for ((typ, f, val), orig) in zip(
                    testing_params.call_args_validators, orig_call_args
                ):
                    self.assertIsInstance(orig, typ)
                    self.assertEqual(f(orig), val)

    def _get_params(self) -> List[PermutedArgspecTestingParams]:
        in_ = MagicMock(spec=POINTER(IUnknown))()
        out = POINTER(IUnknown)()
        inout1 = 5
        inout2 = "abc"
        IN_ARGSPEC = (["in"], POINTER(comtypes.IUnknown), "in")
        OUT_ARGSPEC = (["out"], POINTER(POINTER(comtypes.IUnknown)), "out")
        INOUT1_ARGSPEC = (["in", "out"], POINTER(c_ulong), "inout1")
        INOUT2_ARGSPEC = (["in", "out"], POINTER(WSTRING), "inout2")
        IN_VALIDATOR = (MagicMock, lambda x: x, in_)
        INOUT1_VALIDATOR = (c_ulong, lambda x: x.value, inout1)
        INOUT2_VALIDATOR = (WSTRING, lambda x: x.value, inout2)
        return [
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, OUT_ARGSPEC, INOUT1_ARGSPEC, INOUT2_ARGSPEC),
                (in_, inout1, inout2),
                (out, ..., ...),
                [out, inout1, inout2],
                [IN_VALIDATOR, INOUT1_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, OUT_ARGSPEC, INOUT2_ARGSPEC, INOUT1_ARGSPEC),
                (in_, inout2, inout1),
                (out, ..., ...),
                [out, inout2, inout1],
                [IN_VALIDATOR, INOUT2_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, INOUT1_ARGSPEC, OUT_ARGSPEC, INOUT2_ARGSPEC),
                (in_, inout1, inout2),
                (..., out, ...),
                [inout1, out, inout2],
                [IN_VALIDATOR, INOUT1_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, INOUT1_ARGSPEC, INOUT2_ARGSPEC, OUT_ARGSPEC),
                (in_, inout1, inout2),
                (..., ..., out),
                [inout1, inout2, out],
                [IN_VALIDATOR, INOUT1_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, INOUT2_ARGSPEC, OUT_ARGSPEC, INOUT1_ARGSPEC),
                (in_, inout2, inout1),
                (..., out, ...),
                [inout2, out, inout1],
                [IN_VALIDATOR, INOUT2_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (IN_ARGSPEC, INOUT2_ARGSPEC, INOUT1_ARGSPEC, OUT_ARGSPEC),
                (in_, inout2, inout1),
                (..., ..., out),
                [inout2, inout1, out],
                [IN_VALIDATOR, INOUT2_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, IN_ARGSPEC, INOUT1_ARGSPEC, INOUT2_ARGSPEC),
                (in_, inout1, inout2),
                (out, ..., ...),
                [out, inout1, inout2],
                [IN_VALIDATOR, INOUT1_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, IN_ARGSPEC, INOUT2_ARGSPEC, INOUT1_ARGSPEC),
                (in_, inout2, inout1),
                (out, ..., ...),
                [out, inout2, inout1],
                [IN_VALIDATOR, INOUT2_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, INOUT1_ARGSPEC, IN_ARGSPEC, INOUT2_ARGSPEC),
                (inout1, in_, inout2),
                (out, ..., ...),
                [out, inout1, inout2],
                [INOUT1_VALIDATOR, IN_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, INOUT1_ARGSPEC, INOUT2_ARGSPEC, IN_ARGSPEC),
                (inout1, inout2, in_),
                (out, ..., ...),
                [out, inout1, inout2],
                [INOUT1_VALIDATOR, INOUT2_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, INOUT2_ARGSPEC, IN_ARGSPEC, INOUT1_ARGSPEC),
                (inout2, in_, inout1),
                (out, ..., ...),
                [out, inout2, inout1],
                [INOUT2_VALIDATOR, IN_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (OUT_ARGSPEC, INOUT2_ARGSPEC, INOUT1_ARGSPEC, IN_ARGSPEC),
                (inout2, inout1, in_),
                (out, ..., ...),
                [out, inout2, inout1],
                [INOUT2_VALIDATOR, INOUT1_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, IN_ARGSPEC, OUT_ARGSPEC, INOUT2_ARGSPEC),
                (inout1, in_, inout2),
                (..., out, ...),
                [inout1, out, inout2],
                [INOUT1_VALIDATOR, IN_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, IN_ARGSPEC, INOUT2_ARGSPEC, OUT_ARGSPEC),
                (inout1, in_, inout2),
                (..., ..., out),
                [inout1, inout2, out],
                [INOUT1_VALIDATOR, IN_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, OUT_ARGSPEC, IN_ARGSPEC, INOUT2_ARGSPEC),
                (inout1, in_, inout2),
                (..., out, ...),
                [inout1, out, inout2],
                [INOUT1_VALIDATOR, IN_VALIDATOR, INOUT2_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, OUT_ARGSPEC, INOUT2_ARGSPEC, IN_ARGSPEC),
                (inout1, inout2, in_),
                (..., out, ...),
                [inout1, out, inout2],
                [INOUT1_VALIDATOR, INOUT2_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, INOUT2_ARGSPEC, IN_ARGSPEC, OUT_ARGSPEC),
                (inout1, inout2, in_),
                (..., ..., out),
                [inout1, inout2, out],
                [INOUT1_VALIDATOR, INOUT2_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT1_ARGSPEC, INOUT2_ARGSPEC, OUT_ARGSPEC, IN_ARGSPEC),
                (inout1, inout2, in_),
                (..., ..., out),
                [inout1, inout2, out],
                [INOUT1_VALIDATOR, INOUT2_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, IN_ARGSPEC, OUT_ARGSPEC, INOUT1_ARGSPEC),
                (inout2, in_, inout1),
                (..., out, ...),
                [inout2, out, inout1],
                [INOUT2_VALIDATOR, IN_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, IN_ARGSPEC, INOUT1_ARGSPEC, OUT_ARGSPEC),
                (inout2, in_, inout1),
                (..., ..., out),
                [inout2, inout1, out],
                [INOUT2_VALIDATOR, IN_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, OUT_ARGSPEC, IN_ARGSPEC, INOUT1_ARGSPEC),
                (inout2, in_, inout1),
                (..., out, ...),
                [inout2, out, inout1],
                [INOUT2_VALIDATOR, IN_VALIDATOR, INOUT1_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, OUT_ARGSPEC, INOUT1_ARGSPEC, IN_ARGSPEC),
                (inout2, inout1, in_),
                (..., out, ...),
                [inout2, out, inout1],
                [INOUT2_VALIDATOR, INOUT1_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, INOUT1_ARGSPEC, IN_ARGSPEC, OUT_ARGSPEC),
                (inout2, inout1, in_),
                (..., ..., out),
                [inout2, inout1, out],
                [INOUT2_VALIDATOR, INOUT1_VALIDATOR, IN_VALIDATOR],
            ),
            PermutedArgspecTestingParams(
                (INOUT2_ARGSPEC, INOUT1_ARGSPEC, OUT_ARGSPEC, IN_ARGSPEC),
                (inout2, inout1, in_),
                (..., ..., out),
                [inout2, inout1, out],
                [INOUT2_VALIDATOR, INOUT1_VALIDATOR, IN_VALIDATOR],
            ),
        ]


class Test_Error(ut.TestCase):
    def test_missing_direction(self):
        self_ = MagicMock(name="Self")
        spec = comtypes.COMMETHOD(
            [],
            HRESULT,
            "Foo",
            ([], POINTER(c_ulong)),
        )
        orig = MagicMock(__name__="orig")
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        with self.assertRaises(Exception) as cm:
            fixed(self_, 4)
        self.assertEqual(
            str(cm.exception),
            "A parameter for orig has neither 'out' nor 'in' specified",
        )

    def test_missing_name_omitted(self):
        self_ = MagicMock(name="Self")
        spec = comtypes.COMMETHOD(
            [],
            HRESULT,
            "Foo",
            (["in", "out"], POINTER(c_ulong)),
        )
        orig = MagicMock()
        fixed = _fix_inout_args(orig, spec.argtypes, spec.paramflags)
        with self.assertRaises(Exception) as cm:
            fixed(self_)
        self.assertEqual(
            str(cm.exception), "Unnamed inout parameters cannot be omitted"
        )


if __name__ == "__main__":
    ut.main()
