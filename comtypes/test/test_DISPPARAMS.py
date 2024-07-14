import unittest as ut


class TestCase(ut.TestCase):
    def test(self):
        from comtypes.automation import DISPPARAMS, VARIANT

        dp = DISPPARAMS()
        dp.rgvarg = (VARIANT * 3)()

        for i in range(3):
            self.assertEqual(dp.rgvarg[i].value, None)

        dp.rgvarg[0].value = 42
        dp.rgvarg[1].value = "spam"
        dp.rgvarg[2].value = "foo"

        # damn, there's still this old bug!

        self.assertEqual(dp.rgvarg[0].value, 42)
        # these fail:
        # self.failUnlessEqual(dp.rgvarg[1].value, "spam")
        # self.failUnlessEqual(dp.rgvarg[2].value, "foo")

    def X_test_2(self):
        # basically the same test as above
        from comtypes.automation import DISPPARAMS, VARIANT

        args = [42, None, "foo"]

        dp = DISPPARAMS()
        dp.rgvarg = (VARIANT * 3)(*list(map(VARIANT, args[::-1])))

        import gc

        gc.collect()

        self.assertEqual(dp.rgvarg[0].value, 42)
        self.assertEqual(dp.rgvarg[1].value, "spam")
        self.assertEqual(dp.rgvarg[2].value, "foo")


class Test_DispParamsGenerator(ut.TestCase):
    def _get_rgvargs(self, dp):
        return [dp.rgvarg[i].value for i in range(dp.cArgs)]

    def test_invkind(self):
        from comtypes.automation import (
            DispParamsGenerator,
            DISPATCH_METHOD,
            DISPATCH_PROPERTYGET,
            DISPATCH_PROPERTYPUT,
            DISPATCH_PROPERTYPUTREF,
            DISPID_PROPERTYPUT,
        )

        def _is_null(d):
            return not bool(d)

        def _is_dispid_propput(d):
            return d.contents.value == DISPID_PROPERTYPUT

        for invkind, id_validator, c_namedargs in [
            (DISPATCH_METHOD, lambda d: not bool(d), 0),
            (DISPATCH_PROPERTYGET, _is_null, 0),
            (DISPATCH_PROPERTYPUT, _is_dispid_propput, 1),
            (DISPATCH_PROPERTYPUTREF, _is_dispid_propput, 1),
        ]:
            with self.subTest(
                invkind=invkind, id_validator=id_validator, c_namedargs=c_namedargs
            ):
                dp = DispParamsGenerator(invkind, ()).generate(9)
                self.assertEqual(self._get_rgvargs(dp), [9])
                self.assertTrue(id_validator(dp.rgdispidNamedArgs))
                self.assertEqual(dp.cArgs, 1)
                self.assertEqual(dp.cNamedArgs, c_namedargs)

    def test_c_args(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        for args, c_args in [
            ((), 0),
            ((9,), 1),
            (("foo", 3.14), 2),
            ((2, "bar", 1.41), 3),
        ]:
            with self.subTest(args=args, c_args=c_args):
                dp = DispParamsGenerator(DISPATCH_METHOD, ()).generate(*args)
                self.assertEqual(dp.cArgs, c_args)

    def test_no_argspec(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        gen = DispParamsGenerator(DISPATCH_METHOD, ())
        for args, rgvargs in [
            ((), []),
            ((9,), [9]),
            (("foo", 3.14), [3.14, "foo"]),
            ((2, "bar", 1.41), [1.41, "bar", 2]),
        ]:
            with self.subTest(args=args, rgvargs=rgvargs):
                dp = gen.generate(*args)
                self.assertEqual(self._get_rgvargs(dp), rgvargs)
        with self.assertRaises(TypeError) as ce:
            gen.generate(4, 3.14, "foo", a="spam")
        self.assertEqual(ce.exception.args, ("got an unexpected keyword argument 'a'",))

    def test_argspec_in_x2(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        IN = ["in"]
        spec = ((IN, ..., "a"), (IN, ..., "b"))
        gen = DispParamsGenerator(DISPATCH_METHOD, spec)  # type: ignore
        self.assertEqual(self._get_rgvargs(gen.generate(3, 2.2)), [2.2, 3])
        self.assertEqual(self._get_rgvargs(gen.generate(b=1.4, a=2)), [1.4, 2])
        self.assertEqual(self._get_rgvargs(gen.generate(5, b=3.1)), [3.1, 5])
        with self.assertRaises(TypeError) as ce:
            gen.generate(1, a=2)
        self.assertEqual(ce.exception.args, ("got multiple values for argument 'a'",))
        with self.assertRaises(TypeError) as ce:
            gen.generate(a=3)
        self.assertEqual(
            ce.exception.args, ("missing 1 required positional argument: 'b'",)
        )
        with self.assertRaises(TypeError) as ce:
            gen.generate(a=1, b=2, c=3)
        self.assertEqual(ce.exception.args, ("got an unexpected keyword argument 'c'",))
        # THOSE MIGHT RAISE `COMError` IN CALLER.
        self.assertEqual(self._get_rgvargs(gen.generate()), [])
        self.assertEqual(self._get_rgvargs(gen.generate(1, 2, 3)), [3, 2, 1])

    def test_argspec_in_x1_and_optin_x1(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        IN, OPT_IN = ["in"], ["in", "optional"]
        spec = ((IN, ..., "a"), (OPT_IN, ..., "b", "foo"))
        gen = DispParamsGenerator(DISPATCH_METHOD, spec)  # type: ignore
        self.assertEqual(self._get_rgvargs(gen.generate(2)), [2])
        self.assertEqual(self._get_rgvargs(gen.generate(4, "bar")), ["bar", 4])
        with self.assertRaises(TypeError) as ce:
            gen.generate(b="baz")
        self.assertEqual(
            ce.exception.args, ("missing 1 required positional argument: 'a'",)
        )
        with self.assertRaises(TypeError) as ce:
            gen.generate(4, "bar", b="baz")
        self.assertEqual(ce.exception.args, ("got multiple values for argument 'b'",))

    def test_argspec_in_x3(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        IN = ["in"]
        spec = ((IN, ..., "a"), (IN, ..., "b"), (IN, ..., "c"), (IN, ..., "d"))
        gen = DispParamsGenerator(DISPATCH_METHOD, spec)  # type: ignore
        self.assertEqual(self._get_rgvargs(gen.generate(1, 2, 3, 4)), [4, 3, 2, 1])
        with self.assertRaises(TypeError) as ce:
            gen.generate(d=4)
        self.assertEqual(
            ce.exception.args,
            ("missing 3 required positional arguments: 'a', 'b' and 'c'",),
        )
        with self.assertRaises(TypeError) as ce:
            gen.generate(a=1, c=3)
        self.assertEqual(
            ce.exception.args,
            ("missing 2 required positional arguments: 'b' and 'd'",),
        )
        with self.assertRaises(TypeError) as ce:
            gen.generate(1, b=2, d=4)
        self.assertEqual(
            ce.exception.args,
            ("missing 1 required positional argument: 'c'",),
        )

    def test_argspec_optin_x3(self):
        from comtypes.automation import DispParamsGenerator, DISPATCH_METHOD

        OPT_IN = ["in", "optional"]
        spec = (
            (OPT_IN, ..., "a", 1),
            (OPT_IN, ..., "b", 3.1),
            (OPT_IN, ..., "c", "foo"),
        )
        gen = DispParamsGenerator(DISPATCH_METHOD, spec)  # type: ignore
        self.assertEqual(self._get_rgvargs(gen.generate()), [])
        self.assertEqual(self._get_rgvargs(gen.generate(a=2)), [2])
        self.assertEqual(self._get_rgvargs(gen.generate(b=1.7)), [1.7, 1])
        self.assertEqual(self._get_rgvargs(gen.generate(c="bar")), ["bar", 3.1, 1])
        self.assertEqual(self._get_rgvargs(gen.generate(2, c="bar")), ["bar", 3.1, 2])
        with self.assertRaises(TypeError) as ce:
            gen.generate(d=5)
        self.assertEqual(ce.exception.args, ("got an unexpected keyword argument 'd'",))
        with self.assertRaises(TypeError) as ce:
            gen.generate(3, a=5)
        self.assertEqual(ce.exception.args, ("got multiple values for argument 'a'",))


if __name__ == "__main__":
    ut.main()
