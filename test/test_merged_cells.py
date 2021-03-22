# standard lib imports
from unittest import TestCase
import os, sys

# 3rd party lib support
if sys.version_info[0] == 3:
    from pathlib import Path
else:
    from pathlib2 import Path

# local lib imports
from pylightxl import pylightxl as xl

if 'test' in os.listdir('.'):
    # running from top level
    os.chdir('./test')
DB = xl.readxl('./merged_cells.xlsx')







class TestMergedCells(TestCase):
    def common_tst(self, sheet_name,merged_cells):
        sheet = DB.ws(sheet_name)
        merged_cells_reported = sheet.merged_cells
        self.assertEqual(len(merged_cells), len(merged_cells_reported))
        self.assertEqual(set(merged_cells_reported), set(merged_cells.keys()))
        for k,v in merged_cells.items():
            rlo,rhi,clo,chi = k
            self.assertEqual(sheet.index(rlo,clo),"{}-{}".format(sheet_name, v))

    def test_ws1(self):
        sheet_name = "Sheet1"
        merged_cells = {(1, 21 ,2  ,2) :"B1:B21",
                        (6, 6 ,4  ,10):"D6:J6",
                        (7, 21,5  ,7) :"E7:G21",
                        (3, 33,16 ,19):"P3:S33"}
        self.common_tst(sheet_name, merged_cells)

    def test_ws2(self):
        sheet_name = "Sheet2"
        merged_cells = {(9,48, 6,9):"F9:I48",
                        (6,11,12,17):"L6:Q11"}
        self.common_tst(sheet_name, merged_cells)

    def test_ws3(self):
        sheet_name = "Sheet3"
        merged_cells = {}
        self.common_tst(sheet_name, merged_cells)