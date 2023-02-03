from typing import Any, List, NamedTuple, Optional, Sequence, Tuple, Type
from ctypes import POINTER, c_bool, c_ulong, c_wchar_p
from itertools import permutations
import unittest as ut
from unittest.mock import MagicMock

import comtypes
from comtypes.client import IUnknown
from comtypes._memberspec import _fix_inout_args, _ParamFlagType


class Param(NamedTuple):
    argtype: Type
    paramflags: _ParamFlagType


class TestEntry:
    def __init__(self, test_case: ut.TestCase, param_spec: Sequence[Param]):
        self.test_case = test_case
        self.param_spec = param_spec

    def run_test_with_args(self, *args, **kwargs) -> Tuple[Any, MagicMock]:
        """Runs the test with the provided arguments."""
        inner_mock = MagicMock()

        def mock_function(_, *args, **kwargs):
            # Call the mock
            inner_mock(*args, **kwargs)
            # _fix_inout_args crashes if we don't return the correct types of values.
            # Here we return the inout values unmodified, and a generated value for every purely-out parameter.
            results = []
            arg_stack = list(reversed(args))

            def next_arg(name: Optional[str]):
                """Get the next positional argument if any are left, else try to get the matching keyword argument"""
                nonlocal arg_stack, kwargs
                if arg_stack:
                    return arg_stack.pop()
                elif name is not None:
                    return kwargs.pop(name)
                else:
                    raise TypeError(
                        "Mock: Missing positional argument for nameless parameter"
                    )

            try:
                for argtype, param in self.param_spec:
                    if param[0] & 3 == 3:  # inout
                        results.append(next_arg(param[1]))
                    elif param[0] & 2 == 2:  # out
                        results.append(argtype())
                    elif param[0] & 1 == 1:  # in
                        next_arg(param[1])
                    else:
                        raise ValueError(
                            "Mock: Parameter has neither 'out' nor 'in' specified"
                        )
            except (IndexError, KeyError) as e:
                raise ValueError(f"Mock: Not enough arguments supplied: {e}")

            # Verify that all provided arguments have been used
            if len(arg_stack) > 0 or len(kwargs) > 0:
                raise ValueError("Mock: Too many arguments supplied")

            if len(results) == 0:
                return None
            elif len(results) == 1:
                return results[0]
            else:
                return tuple(results)

        argtypes = tuple(x.argtype for x in self.param_spec)
        paramflags = tuple(x.paramflags for x in self.param_spec)
        fixed_fn = comtypes.instancemethod(
            _fix_inout_args(mock_function, argtypes, paramflags), self, None
        )

        result = fixed_fn(*args, **kwargs)

        return (result, inner_mock)

    def run_test(self):
        """Runs the test with automatically generated positional arguments"""
        args = [x.argtype() for x in self.param_spec if x.paramflags[0] & 1 == 1]
        results, mock = self.run_test_with_args(*args)
        mock.assert_called_once_with(*args)
        out_params = [x for x in self.param_spec if x.paramflags[0] & 2 == 2]
        if len(out_params) == 0:
            self.test_case.assertIsNone(results)
        elif len(out_params) == 1:
            self.test_case.assertIsInstance(results, out_params[0].argtype)
        else:
            self.test_case.assertEqual(len(results), len(out_params))
            for result, param in zip(results, out_params):
                self.test_case.assertIsInstance(result, param.argtype)


class Test_InOut_args(ut.TestCase):
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
        self.assertEqual(
            cm.exception.args[0],
            "A parameter for mock_function has neither 'out' nor 'in' specified",
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
        default_ulong = c_ulong()
        self.assertEqual(result, default_ulong.value)
        mock.assert_called_once()
        generated_arg = mock.call_args[1]["param_name"]
        self.assertIsInstance(generated_arg, c_ulong)
        self.assertEqual(generated_arg.value, default_ulong.value)

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

        internal_kwargs: dict[str, Any] = mock.call_args[1]
        self.assertEqual(len(internal_kwargs), 3)
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

    # TODO I didn't find a natural case to test the following lines:
    #
    # else:
    #   v = atyp.from_param(v)
    #   assert not isinstance(v, BYREFTYPE)
    #
    # We might be able to construct a subclass of IUnknown, override its .from_param method,
    # and then accept e.g. an integer as an argument. However, this feels rather articifial.


if __name__ == "__main__":
    ut.main()
