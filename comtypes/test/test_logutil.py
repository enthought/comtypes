import unittest as ut

from comtypes.logutil import deprecated


class Test_deprecated(ut.TestCase):
    def test_warning_is_raised(self):
        reason_text = "This is deprecated."

        @deprecated(reason_text)
        def test_func():
            return "success"

        with self.assertWarns(DeprecationWarning) as cm:
            result = test_func()
        self.assertEqual(result, "success")
        self.assertEqual(reason_text, str(cm.warning))
