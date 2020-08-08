# standard lib imports
from unittest import TestCase

from pylightxl.pylightxl import Database, Worksheet


class TestWorksheet(TestCase):

    def test_update_index(self):
        ws = Worksheet({})
        ws.update_index(row=4, col=2, val=42)
        self.assertEqual(ws.size, [4, 2])
        self.assertEqual(ws.index(4, 2), 42)
        self.assertEqual(ws.address('B4'), 42)
        self.assertEqual(ws.row(4)[1], 42)
        self.assertEqual(ws.col(2)[3], 42)