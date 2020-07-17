# standard lib imports
from unittest import TestCase

# local lib imports
from pylightxl.readxl import readxl
from pylightxl.database import Database, Worksheet, address2index, index2address, \
    columnletter2num, num2columnletters

try:
    # running from top level
    DB = readxl('./test/testbook.xlsx')
except ValueError:
    # running within /test folder or in debug
    DB = readxl('./testbook.xlsx')


# TODO: add test for formula returning empty string </v> tag

class test_readxl_bad_input(TestCase):

    def test_bad_fn_type(self):
        with self.assertRaises(ValueError) as e:
            db = readxl(fn=1)
            self.assertEqual(e, 'Error - Incorrect file entry ({}).'.format('1'))

    def test_bad_fn_exist(self):
        with self.assertRaises(ValueError) as e:
            db = readxl('bad')
            self.assertEqual(e, 'Error - File ({}) does not exit.'.format('bad'))

    def test_bad_fn_ext(self):
        with self.assertRaises(ValueError) as e:
            try:
                db = readxl('test/test_readxl.py')
            except ValueError:
                db = readxl('test_readxl.py')
            self.assertEqual(e, 'Error - Incorrect Excel file extension ({}). '
                                'File extension supported: .xlsx .xlsm'.format('py'))


class test_readxl_integration(TestCase):

    def test_AllSheetsRead(self):
        db_ws_names = DB.ws_names
        db_ws_names.sort()
        # test 10+ sheets to ensure sorting matches correctly
        true_ws_names = ['empty', 'types', 'scatter', 'merged_cells', 'length', 'sheet_not_to_read',
                         'ssd_error1', 'ssd_error2', 'ssd_error3', 'semistrucdata1', 'semistrucdata2']
        true_ws_names.sort()
        self.assertEqual(db_ws_names, true_ws_names)

    def test_SelectedSheetReading(self):
        try:
            db = readxl('test/testbook.xlsx', ('empty', 'types'))
        except ValueError:
            db = readxl('testbook.xlsx', ('empty', 'types'))
        db_ws_names = db.ws_names
        db_ws_names.sort()
        true_ws_names = ['empty', 'types']
        true_ws_names.sort()
        self.assertEqual(db_ws_names, true_ws_names)

    def test_commondString(self):
        # all cells that contain strings (without equations are stored in a commondString.xlm)
        self.assertEqual(DB.ws('types').address('A2'), 'copy')
        # leading space comes out different in xml; <t xlm:space="preserve">
        self.assertEqual(DB.ws('types').address('B3'), ' leadingspace')
        self.assertEqual(DB.ws('types').address('B4'), 'copy')

    def test_ws_empty(self):
        # should not contain any cell data, however the user should be able to index to any cell for ""
        self.assertEqual(DB.ws('empty').index(1, 1), '')
        self.assertEqual(DB.ws('empty').index(10, 10), '')
        self.assertEqual(DB.ws('empty').size, [0, 0])
        self.assertEqual(DB.ws('empty').row(1), [])
        self.assertEqual(DB.ws('empty').col(1), [])

    def test_ws_types(self):
        self.assertEqual(DB.ws('types').index(1, 1), 11)
        self.assertEqual(DB.ws('types').index(2, 1), 'copy')
        self.assertEqual(DB.ws('types').index(3, 1), 31)
        self.assertEqual(DB.ws('types').index(4, 1), 41)
        self.assertEqual(DB.ws('types').index(5, 1), 'string from A2 copy')
        self.assertEqual(DB.ws('types').index(6, 1), '')

        self.assertEqual(DB.ws('types').index(1, 2), 12.1)
        self.assertEqual(DB.ws('types').index(2, 2), '"22"')
        self.assertEqual(DB.ws('types').index(3, 2), ' leadingspace')
        self.assertEqual(DB.ws('types').index(4, 2), 'copy')
        self.assertEqual(DB.ws('types').index(5, 2), '')

        self.assertEqual(DB.ws('types').index(1, 3), '')
        self.assertEqual(DB.ws('types').size, [5, 2])

        self.assertEqual(DB.ws('types').row(1), [11, 12.1])
        self.assertEqual(DB.ws('types').row(2), ['copy', '"22"'])
        self.assertEqual(DB.ws('types').row(3), [31, ' leadingspace'])
        self.assertEqual(DB.ws('types').row(4), [41, 'copy'])
        self.assertEqual(DB.ws('types').row(5), ['string from A2 copy', ''])
        self.assertEqual(DB.ws('types').row(6), ['', ''])

        self.assertEqual(DB.ws('types').col(1), [11, 'copy', 31, 41, 'string from A2 copy'])
        self.assertEqual(DB.ws('types').col(2), [12.1, '"22"', ' leadingspace', 'copy', ''])
        self.assertEqual(DB.ws('types').col(3), ['', '', '', '', ''])

        for i, row in enumerate(DB.ws('types').rows, start=1):
            self.assertEqual(row, DB.ws('types').row(i))
        for i, col in enumerate(DB.ws('types').cols, start=1):
            self.assertEqual(col, DB.ws('types').col(i))

        self.assertEqual(DB.ws('types').keycol(11), [11, 'copy', 31, 41, 'string from A2 copy'])
        self.assertEqual(DB.ws('types').keyrow(11), [11, 12.1])

    def test_ws_scatter(self):
        self.assertEqual(DB.ws('scatter').index(1, 1), '')
        self.assertEqual(DB.ws('scatter').index(2, 2), 22)
        self.assertEqual(DB.ws('scatter').index(3, 3), 33)
        self.assertEqual(DB.ws('scatter').index(3, 4), 34)
        self.assertEqual(DB.ws('scatter').index(6, 6), 66)
        self.assertEqual(DB.ws('scatter').index(5, 6), '')

        self.assertEqual(DB.ws('scatter').size, [6, 6])

    def test_ws_merged_cells(self):
        self.assertEqual(DB.ws('merged_cells').index(1, 1), '')
        self.assertEqual(DB.ws('merged_cells').index(1, 2), 12)
        self.assertEqual(DB.ws('merged_cells').index(1, 3), 13)
        self.assertEqual(DB.ws('merged_cells').index(2, 1), 21)
        self.assertEqual(DB.ws('merged_cells').index(2, 2), 22)
        self.assertEqual(DB.ws('merged_cells').index(2, 3), 23)
        self.assertEqual(DB.ws('merged_cells').index(3, 2), 32)
        self.assertEqual(DB.ws('merged_cells').index(4, 3), 43)
        self.assertEqual(DB.ws('merged_cells').index(5, 2), 52)
        self.assertEqual(DB.ws('merged_cells').index(6, 1), 61)
        self.assertEqual(DB.ws('merged_cells').index(7, 2), 72)
        self.assertEqual(DB.ws('merged_cells').index(7, 3), 73)
        self.assertEqual(DB.ws('merged_cells').index(8, 3), 83)
        self.assertEqual(DB.ws('merged_cells').index(9, 1), 91)
        self.assertEqual(DB.ws('merged_cells').index(9, 3), 93)
        self.assertEqual(DB.ws('merged_cells').index(10, 3), 103)

    def test_ws_length(self):
        self.assertEqual(DB.ws('length').size, [1048576, 16384])


class test_Database(TestCase):
    db = Database()

    def test_db_badsheet(self):
        db = Database()
        with self.assertRaises(ValueError) as e:
            db.ws('not a sheet')
            self.assertEqual(e, 'Error - Sheetname (not a sheet) is not in the database')

    def test_db_init(self):
        # locally defined to return an empty ws
        db = Database()
        self.assertEqual(db._ws, {})

    def test_db_repr(self):
        self.assertEqual(str(self.db), 'pylightxl.Database')

    def test_db_ws_names(self):
        # locally defined to return an empty list
        db = Database()
        self.assertEqual(db.ws_names, [])

    def test_db_add_ws(self):
        self.db.add_ws(sheetname='test1', data={})
        self.assertEqual(str(self.db.ws(sheetname='test1')), 'pylightxl.Database.Worksheet')
        self.assertEqual(self.db.ws_names, ['test1'])
        self.db.add_ws('test2', {})
        self.assertEqual(self.db.ws_names, ['test1', 'test2'])

    def test_semistrucdata(self):
        table1 = DB.ws('semistrucdata1').ssd()[0]
        table2 = DB.ws('semistrucdata1').ssd()[1]
        table3 = DB.ws('semistrucdata1').ssd()[2]

        table4 = DB.ws('semistrucdata1').ssd(keyrow='myrows', keycol='mycols')[0]

        self.assertEqual(table1, {'keyrows': ['r1', 'r2', 'r3'], 'keycols': ['c1', 'c2'],
                                  'data': [[11, 12], [21, 22], [31, 32]]})
        self.assertEqual(table2, {'keyrows': ['rr1', 'rr2'], 'keycols': ['cc1', 'cc2'],
                                  'data': [[10, 20], [30, 40]]})
        self.assertEqual(table3, {'keyrows': ['rrr1', 'rrr2', 'rrr3'], 'keycols': ['ccc1', 'ccc2'],
                                  'data': [[110, 120], [210, 220], [310, 320]]})

        self.assertEqual(table4, {'keyrows': ['rrrr1'], 'keycols': ['cccc1', 'cccc2', 'cccc3'],
                                  'data': [['one', 'two', 'three']]})

        with self.assertRaises(ValueError) as e:
            _ = DB.ws('semistrucdata2').ssd()
            self.assertEqual(e,
                             'Error - keyrows != keycols most likely due to missing keyword flag '
                             'keycol IDs: [1], keyrow IDs: []')

    def test_new_empty_cell(self):
        self.assertEqual(DB.ws('empty').index(1, 1), '')
        DB.set_emptycell(val='NA')
        self.assertEqual(DB.ws('empty').index(1, 1), 'NA')
        DB.set_emptycell(val=0)
        self.assertEqual(DB.ws('empty').index(1, 1), 0)
        # reset it so other tests run correctly
        DB.set_emptycell(val='')



class test_Worksheet(TestCase):

    def test_ws_init(self):
        ws = Worksheet(data={})
        self.assertEqual(ws._data, {})
        self.assertEqual(ws.maxrow, 0)
        self.assertEqual(ws.maxcol, 0)

    def test_ws_repr(self):
        ws = Worksheet({})
        self.assertEqual(str(ws), 'pylightxl.Database.Worksheet')

    def test_ws_calc_size(self):
        ws = Worksheet({})
        # force calc size
        ws._calc_size()
        self.assertEqual(ws.maxrow, 0)
        self.assertEqual(ws.maxcol, 0)

        ws._data = {'A1': {'v': 11}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1)
        self.assertEqual(ws.maxcol, 1)

        ws._data = {'A1': {'v': 11}, 'A2': {'v': 21}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 1)

        ws._data = {'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 2)

        ws._data = {'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}, 'B2': {'v': 22}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 2)

        ws._data = {'A1': {'v': 1}, 'AA1': {'v': 27}, 'AAA1': {'v': 703}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1)
        self.assertEqual(ws.maxcol, 703)

        ws._data = {'A1': {'v': 1}, 'A1000': {'v': 1000}, 'A1048576': {'v': 1048576}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1048576)
        self.assertEqual(ws.maxcol, 1)

        ws._data = {'A1': {'v': 1}, 'AA1': {'v': 27}, 'AAA1': {'v': 703}, 'XFD1': {'v': 16384},
                    'A1048576': {'v': 1048576}}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1048576)
        self.assertEqual(ws.maxcol, 16384)

    def test_ws_size(self):
        ws = Worksheet({})
        self.assertEqual(ws.size, [0, 0])
        ws._data = {'A1': {'v': 11}, 'A2': {'v': 21}}
        ws._calc_size()
        self.assertEqual(ws.size, [2, 1])

    def test_ws_address(self):
        ws = Worksheet({'A1': {'v': 1}})
        self.assertEqual(ws.address(address='A1'), 1)
        self.assertEqual(ws.address('A2'), '')

    def test_ws_index(self):
        ws = Worksheet({'A1': {'v': 1}})
        self.assertEqual(ws.index(row=1, col=1), 1)
        self.assertEqual(ws.index(1, 2), '')

    def test_ws_row(self):
        ws = Worksheet({'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}})
        self.assertEqual(ws.row(row=1), [11, 12])
        self.assertEqual(ws.row(2), [21, ''])
        self.assertEqual(ws.row(3), ['', ''])

    def test_ws_col(self):
        ws = Worksheet({'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}})
        self.assertEqual(ws.col(col=1), [11, 21])
        self.assertEqual(ws.col(2), [12, ''])
        self.assertEqual(ws.col(3), ['', ''])

    def test_ws_rows(self):
        ws = Worksheet({'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}})
        correct_list = [[11, 12], [21, '']]
        for i, row in enumerate(ws.rows):
            self.assertEqual(row, correct_list[i])

    def test_ws_cols(self):
        ws = Worksheet({'A1': {'v': 11}, 'A2': {'v': 21}, 'B1': {'v': 12}})
        correct_list = [[11, 21], [12, '']]
        for i, col in enumerate(ws.cols):
            self.assertEqual(col, correct_list[i])

    def test_ws_keycol(self):
        ws = Worksheet({'A1': {'v': 11}, 'B1': {'v': 11}, 'C1': {'v': 13},
                        'A2': {'v': 21}, 'B2': {'v': 22}, 'C2': {'v': 23},
                        'A3': {'v': 11}, 'B3': {'v': 32}, 'C3': {'v': 33}})
        self.assertEqual(ws.keycol(key=11), [11, 21, 11])
        self.assertEqual(ws.keycol(key=11, keyindex=1), [11, 21, 11])
        self.assertEqual(ws.keycol(key=11, keyindex=2), [])
        self.assertEqual(ws.keycol(key=32, keyindex=3), [11, 22, 32])

        self.assertEqual(ws.keyrow(key=11), [11, 11, 13])
        self.assertEqual(ws.keyrow(key=11, keyindex=1), [11, 11, 13])
        self.assertEqual(ws.keyrow(key=11, keyindex=2), [11, 11, 13])
        self.assertEqual(ws.keyrow(key=22, keyindex=2), [21, 22, 23])
        self.assertEqual(ws.keyrow(key=22, keyindex=3), [])


class test_conversion(TestCase):

    def test_address2index_baddata(self):
        with self.assertRaises(ValueError) as e:
            address2index(address=1)
            self.assertEqual(e, 'Error - Address (1) must be a string.')

        with self.assertRaises(ValueError) as e:
            address2index('')
            self.assertEqual(e, 'Error - Address ('') cannot be an empty str.')

        with self.assertRaises(ValueError) as e:
            address2index('1')
            self.assertEqual(e, 'Error - Incorrect address (1) entry. Address must be an alphanumeric '
                                'where the starting character(s) are alpha characters a-z')

        with self.assertRaises(ValueError) as e:
            address2index('1A')
            self.assertEqual(e, 'Error - Incorrect address (1A) entry. Address must be an alphanumeric '
                                'where the starting character(s) are alpha characters a-z')

        with self.assertRaises(ValueError) as e:
            address2index('AA')
            self.assertEqual(e, 'Error - Incorrect address (AA) entry. Address must be an alphanumeric '
                                'where the trailing character(s) are numeric characters 1-9')

    def test_address2index(self):
        self.assertEqual(address2index('A1'), [1, 1])
        self.assertEqual(address2index('A1000'), [1000, 1])
        self.assertEqual(address2index('A1048576'), [1048576, 1])

        self.assertEqual(address2index('Z1'), [1, 26])
        self.assertEqual(address2index('AA1'), [1, 27])
        self.assertEqual(address2index('BA1'), [1, 53])
        self.assertEqual(address2index('YQ1'), [1, 667])
        self.assertEqual(address2index('AAA1'), [1, 703])
        self.assertEqual(address2index('PZD1'), [1, 11496])
        self.assertEqual(address2index('QGK1'), [1, 11685])
        self.assertEqual(address2index('XFD1'), [1, 16384])

        self.assertEqual(address2index('XFD1048576'), [1048576, 16384])

    def test_index2address_baddata(self):
        with self.assertRaises(ValueError) as e:
            index2address(row='', col=1)
            self.assertEqual(e, 'Error - Incorrect row ('') entry. Row must either be a int or float')
        with self.assertRaises(ValueError) as e:
            index2address(1, '')
            self.assertEqual(e, 'Error - Incorrect col ('') entry. Col must either be a int or float')
        with self.assertRaises(ValueError) as e:
            index2address(0, 0)
            self.assertEqual(e, 'Error - Row (0) and Col (0) entry cannot be less than 1')

    def test_index2address(self):
        self.assertEqual(index2address(1, 1), 'A1')
        self.assertEqual(index2address(1000, 1), 'A1000')
        self.assertEqual(index2address(1048576, 1), 'A1048576')

        self.assertEqual(index2address(1, 26), 'Z1')
        self.assertEqual(index2address(1, 27), 'AA1')
        self.assertEqual(index2address(1, 53), 'BA1')
        self.assertEqual(index2address(1, 667), 'YQ1')
        self.assertEqual(index2address(1, 703), 'AAA1')
        self.assertEqual(index2address(1, 11496), 'PZD1')
        self.assertEqual(index2address(1, 11685), 'QGK1')
        self.assertEqual(index2address(1, 16384), 'XFD1')

        self.assertEqual(index2address(1048576, 16384), 'XFD1048576')

    def test_col2num(self):
        self.assertEqual(columnletter2num('A'), 1)
        self.assertEqual(columnletter2num('Z'), 26)
        self.assertEqual(columnletter2num('AA'), 27)
        self.assertEqual(columnletter2num('BA'), 53)
        self.assertEqual(columnletter2num('YQ'), 667)
        self.assertEqual(columnletter2num('ZZ'), 702)
        self.assertEqual(columnletter2num('AAA'), 703)
        self.assertEqual(columnletter2num('PZD'), 11496)
        self.assertEqual(columnletter2num('QGK'), 11685)
        self.assertEqual(columnletter2num('XFD'), 16384)

    def test_num2col(self):
        self.assertEqual(num2columnletters(1), 'A')
        self.assertEqual(num2columnletters(26), 'Z')
        self.assertEqual(num2columnletters(27), 'AA')
        self.assertEqual(num2columnletters(53), 'BA')
        self.assertEqual(num2columnletters(667), 'YQ')
        self.assertEqual(num2columnletters(702), 'ZZ')
        self.assertEqual(num2columnletters(703), 'AAA')
        self.assertEqual(num2columnletters(11496), 'PZD')
        self.assertEqual(num2columnletters(11685), 'QGK')
        self.assertEqual(num2columnletters(16384), 'XFD')
