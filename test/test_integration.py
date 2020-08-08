# standard lib imports
from unittest import TestCase
from os import remove

# local lib imports
from pylightxl.pylightxl import readxl
from pylightxl.pylightxl import writexl
from pylightxl.pylightxl import Database, Worksheet


class TestIntegration(TestCase):

    def test_reading_written_ws(self):
        file_path = 'temporary_test_file.xlsx'
        db = Database()
        db.add_ws('new_ws', {})
        writexl(db, file_path)
        db = readxl(file_path)
        self.assertEqual(db.ws_names, ['new_ws'])
        remove(file_path)

    def test_reading_written_cells(self):
        file_path = 'temporary_test_file.xlsx'
        db = Database()
        db.add_ws('new_ws', {})
        ws = db.ws('new_ws')
        ws.update_index(row=4, col=2, val=42)
        writexl(db, file_path)
        db = readxl(file_path)
        self.assertEqual(db.ws('new_ws').index(4, 2), 42)
        remove(file_path)
