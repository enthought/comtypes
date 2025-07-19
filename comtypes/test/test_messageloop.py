import unittest
from unittest.mock import Mock

from comtypes.messageloop import _MessageLoop


class InsertAndRemoveFilterTest(unittest.TestCase):
    def test_insert_and_remove_filter(self):
        msgloop = _MessageLoop()
        flt = Mock(return_value=[False])
        msgloop.insert_filter(flt)
        msgloop.remove_filter(flt)
