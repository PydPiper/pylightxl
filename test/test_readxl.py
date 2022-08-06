# standard lib imports
from unittest import TestCase
import os, sys

# 3rd party lib support

if sys.version_info[0] == 3:
    from pathlib import Path
else:
    from pathlib2 import Path
    ModuleNotFoundError = ImportError

# local lib imports
try:
    from pylightxl import pylightxl as xl
except ModuleNotFoundError:
    sys.path.append('..')
    from pylightxl import pylightxl as xl

if 'test' in os.listdir('.'):
    # running from top level
    os.chdir('./test')
DB = xl.readxl('./testbook.xlsx')


class TestReadxl_BadInput(TestCase):

    def test_bad_fn_type(self):
        with self.assertRaises(UserWarning) as e:
            _ = xl.readxl(fn=1)
            self.assertEqual('pylightxl - Incorrect file entry ({}).'.format('1'), e)

    def test_bad_fn_exist(self):
        with self.assertRaises(UserWarning) as e:
            _ = xl.readxl('bad')
            self.assertEqual('pylightxl - File ({}) does not exist.'.format('bad'), e)

    def test_bad_fn_ext(self):
        with self.assertRaises(UserWarning) as e:
            _ = xl.readxl('test_read.py')
            self.assertEqual('pylightxl - Incorrect Excel file extension ({}). '
                             'File extension supported: .xlsx .xlsm'.format('py'), e)

    def test_bad_readxl_sheetnames(self):
        with self.assertRaises(UserWarning) as e:
            _ = xl.readxl(fn='./testbook.xlsx', ws='not-a-sheet')
            self.assertRaises('pylightxl - Sheetname ({}) is not in the workbook.'.format('not-a-sheet'), e)

    def test_bad_readxl_extension(self):
        with self.assertRaises(UserWarning) as e:
            _ = xl.readxl(fn='./input.csv')
            self.assertRaises('pylightxl - Incorrect Excel file extension ({}). '
                              'File extension supported: .xlsx .xlsm'.format('csv'), e)

    def test_bad_readxl_workbook_format(self):
        msg = ('pylightxl - Ill formatted workbook.xml. '
               'Skipping NamedRange not containing sheet reference (ex: "Sheet1!A1"): '
               '{name} - {fulladdress}'.format(name='single_nr', fulladdress='$E$6'))
        if sys.version_info[0] > 2:
            # assertWarns only available in py 3.2+
            with self.assertWarns(UserWarning, msg=msg):
                _ = xl.readxl_get_workbook('./bad_nr_workbook.zip')


class TestReadCSV(TestCase):

    def test_readcsv(self):
        db = xl.readcsv(fn='input.csv', delimiter='\t', ws='sh2')

        self.assertEqual(11, db.ws('sh2').index(1, 1))
        self.assertEqual(12.0, db.ws('sh2').index(1, 2))
        self.assertEqual(0.13, db.ws('sh2').index(1, 3))
        self.assertEqual("'14'", db.ws('sh2').index(1, 4))
        self.assertEqual(" ", db.ws('sh2').index(1, 5))
        self.assertEqual(16, db.ws('sh2').index(1, 6))
        self.assertEqual('', db.ws('sh2').index(2, 1))
        self.assertEqual('', db.ws('sh2').index(2, 2))
        self.assertEqual('', db.ws('sh2').index(2, 3))
        self.assertEqual('', db.ws('sh2').index(2, 4))
        self.assertEqual('', db.ws('sh2').index(2, 5))
        self.assertEqual('', db.ws('sh2').index(2, 6))

        self.assertEqual(31, db.ws('sh2').index(4, 1))
        self.assertEqual('', db.ws('sh2').index(4, 2))
        self.assertEqual(False, db.ws('sh2').index(4, 3))
        self.assertEqual('', db.ws('sh2').index(4, 4))
        self.assertEqual(True, db.ws('sh2').index(4, 5))
        self.assertEqual('', db.ws('sh2').index(4, 6))
        self.assertEqual(42, db.ws('sh2').index(5, 2))
        self.assertEqual(' ', db.ws('sh2').index(5, 4))

        self.assertEqual([5, 6], db.ws('sh2').size)


class TestIntegration(TestCase):

    def test_filehandle_readxl(self):
        mypath = Path('./testbook.xlsx')
        if sys.version_info[0] == 3:
            with open(mypath, 'rb') as f:
                db = xl.readxl(fn=f, ws=['types', ])
            self.assertEqual(11, db.ws('types').index(1, 1))

    def test_pathlib_readxl(self):
        mypath = Path('./testbook.xlsx')

        db = xl.readxl(fn=mypath, ws=['types', ])
        self.assertEqual(11, db.ws('types').index(1, 1))

    def test_pathlib_readcsv(self):
        mypath = Path('./input.csv')

        db = xl.readcsv(fn=mypath, delimiter='\t', ws='sh1')
        self.assertEqual(11, db.ws('sh1').index(1, 1))

    def test_AllSheetsRead(self):
        db_ws_names = DB.ws_names
        true_ws_names = ['empty', 'types', 'scatter', 'merged_cells', 'length', 'sheet_not_to_read',
                         'ssd_error1', 'ssd_error2', 'ssd_error3', 'semistrucdata1', 'semistrucdata2']
        self.assertEqual(sorted(true_ws_names), sorted(db_ws_names))

    def test_SelectedSheetReading(self):
        db = xl.readxl('testbook.xlsx', ('empty', 'types'))
        db_ws_names = db.ws_names
        db_ws_names.sort()
        true_ws_names = ['empty', 'types']
        true_ws_names.sort()
        self.assertEqual(true_ws_names, db_ws_names)

    def test_ReadFileStream(self):
        with open('testbook.xlsx', 'rb') as f:
            db = xl.readxl(f, ('empty', 'types'))
        db_ws_names = db.ws_names
        db_ws_names.sort()
        true_ws_names = ['empty', 'types']
        true_ws_names.sort()
        self.assertEqual(true_ws_names, db_ws_names)

    def test_commondString(self):
        # all cells that contain strings (without equations are stored in a commondString.xlm)
        self.assertEqual('copy', DB.ws('types').address('A2'))
        # leading space comes out different in xml; <t xlm:space="preserve">
        self.assertEqual(' leadingspace', DB.ws('types').address('B3'))
        self.assertEqual('copy', DB.ws('types').address('B4'))

    def test_ws_empty(self):
        # should not contain any cell data, however the user should be able to index to any cell for ""
        self.assertEqual('', DB.ws('empty').index(1, 1))
        self.assertEqual('', DB.ws('empty').index(10, 10))
        self.assertEqual([0, 0], DB.ws('empty').size)
        self.assertEqual([], DB.ws('empty').row(1))
        self.assertEqual([], DB.ws('empty').col(1))

    def test_ws_types(self):
        self.assertEqual(11, DB.ws('types').index(1, 1))
        self.assertEqual('comment1', DB.ws('types').index(1, 1, output='c'))
        self.assertEqual('copy', DB.ws('types').index(2, 1))
        self.assertEqual(31, DB.ws('types').index(3, 1))
        self.assertEqual('=31', DB.ws('types').index(3, 1, output='f'))
        self.assertEqual(41, DB.ws('types').index(4, 1))
        self.assertEqual('=A1+30', DB.ws('types').index(4, 1, output='f'))
        self.assertEqual('string from A2 copy', DB.ws('types').index(5, 1))
        self.assertEqual('="string from A2 "&A2', DB.ws('types').index(5, 1, output='f'))
        self.assertEqual(True, DB.ws('types').index(6, 1))
        self.assertEqual(' true', DB.ws('types').index(7, 1))
        self.assertEqual('2021/04/10', DB.ws('types').index(8, 1))
        self.assertEqual('2021/04/10', DB.ws('types').index(9, 1))
        self.assertEqual('02:48:02', DB.ws('types').index(10, 1))
        self.assertEqual('2021/04/10 05:12:00', DB.ws('types').index(11, 1))
        self.assertEqual('', DB.ws('types').index(12, 1))

        self.assertEqual(12.1, DB.ws('types').index(1, 2))
        self.assertEqual('"22"', DB.ws('types').index(2, 2))
        self.assertEqual(' leadingspace', DB.ws('types').index(3, 2))
        self.assertEqual('copy', DB.ws('types').index(4, 2))
        self.assertEqual('', DB.ws('types').index(5, 2))
        self.assertEqual(False, DB.ws('types').index(6, 2))
        self.assertEqual('"false"', DB.ws('types').index(7, 2))
        self.assertEqual('', DB.ws('types').index(8, 2))


        self.assertEqual(-1, DB.ws('types').index(1, 3))
        self.assertEqual('comment2', DB.ws('types').index(1, 3, output='c'))
        self.assertEqual('', DB.ws('types').index(1, 4))
        self.assertEqual([11, 3], DB.ws('types').size)

        self.assertEqual([11, 12.1, -1], DB.ws('types').row(1))
        self.assertEqual(['copy', '"22"', ''], DB.ws('types').row(2))
        self.assertEqual([31, ' leadingspace', ''], DB.ws('types').row(3))
        self.assertEqual([41, 'copy', ''], DB.ws('types').row(4))
        self.assertEqual(['string from A2 copy', '', ''], DB.ws('types').row(5))
        self.assertEqual([True, False, ''], DB.ws('types').row(6))
        self.assertEqual([' true', '"false"', ''], DB.ws('types').row(7))
        self.assertEqual(['2021/04/10', '', ''], DB.ws('types').row(8))
        self.assertEqual(['2021/04/10', '', ''], DB.ws('types').row(9))
        self.assertEqual(['02:48:02', '', ''], DB.ws('types').row(10))
        self.assertEqual(['2021/04/10 05:12:00', '', ''], DB.ws('types').row(11))
        self.assertEqual(['', '', ''], DB.ws('types').row(12))

        self.assertEqual([11, 'copy', 31, 41, 'string from A2 copy', True, ' true',
                          '2021/04/10', '2021/04/10', '02:48:02', '2021/04/10 05:12:00'], DB.ws('types').col(1))
        self.assertEqual([12.1, '"22"', ' leadingspace', 'copy', '',  False, '"false"', '', '', '', ''], DB.ws('types').col(2))
        self.assertEqual([-1, '', '', '', '', '', '', '', '', '', ''], DB.ws('types').col(3))

        for i, row in enumerate(DB.ws('types').rows, start=1):
            self.assertEqual(DB.ws('types').row(i), row)
        for i, col in enumerate(DB.ws('types').cols, start=1):
            self.assertEqual(DB.ws('types').col(i), col)

        self.assertEqual([11, 'copy', 31, 41, 'string from A2 copy', True, ' true',
                          '2021/04/10', '2021/04/10', '02:48:02', '2021/04/10 05:12:00'], DB.ws('types').keycol(11))
        self.assertEqual([11, 12.1, -1], DB.ws('types').keyrow(11))

    def test_ws_scatter(self):
        self.assertEqual('', DB.ws('scatter').index(1, 1))
        self.assertEqual(22, DB.ws('scatter').index(2, 2))
        self.assertEqual('comment3', DB.ws('scatter').index(2, 2, output='c'))
        self.assertEqual(33, DB.ws('scatter').index(3, 3))
        self.assertEqual(34, DB.ws('scatter').index(3, 4))
        self.assertEqual(66, DB.ws('scatter').index(6, 6))
        self.assertEqual('', DB.ws('scatter').index(5, 6))

        self.assertEqual([6, 6], DB.ws('scatter').size)

    def test_ws_length(self):
        self.assertEqual([1048576, 16384], DB.ws('length').size)

    def test_reading_written_ws(self):
        file_path = 'temporary_test_file.xlsx'
        db = xl.Database()
        db.add_ws('new_ws')
        xl.writexl(db, file_path)
        db = xl.readxl(file_path)
        self.assertEqual(['new_ws'], db.ws_names)
        os.remove(file_path)

    def test_reading_written_cells(self):
        file_path = 'temporary_test_file.xlsx'
        if file_path in os.listdir('.'):
            os.remove(file_path)
        db = xl.Database()
        db.add_ws('new_ws', {})
        ws = db.ws('new_ws')
        ws.update_index(row=4, col=2, val=42)
        xl.writexl(db, file_path)
        db = xl.readxl(file_path)
        self.assertEqual(42, db.ws('new_ws').index(4, 2))
        os.remove(file_path)

    def test_reading_nr(self):
        true_nr = {'table1': 'semistrucdata1!A1:C4',
                   'table2': 'semistrucdata1!G1:I3',
                   'table3': 'semistrucdata1!A11:A14',
                   'single_nr': 'semistrucdata1!E6',
                   }
        self.assertEqual(true_nr, DB.nr_names)

    def test_semistrucdata(self):
        table1 = DB.ws('semistrucdata1').ssd()[0]
        table2 = DB.ws('semistrucdata1').ssd()[1]
        table3 = DB.ws('semistrucdata1').ssd()[2]

        table4 = DB.ws('semistrucdata1').ssd(keyrows='myrows', keycols='mycols')[0]

        self.assertEqual({'keyrows': ['r1', 'r2', 'r3'], 'keycols': ['c1', 'c2'],
                          'data': [[11, 12], [21, 22], [31, 32]]}, table1)
        self.assertEqual({'keyrows': ['rr1', 'rr2'], 'keycols': ['cc1', 'cc2'],
                          'data': [[10, 20], [30, 40]]}, table2)
        self.assertEqual({'keyrows': ['rrr1', 'rrr2', 'rrr3'], 'keycols': ['ccc1', 'ccc2'],
                          'data': [[110, 120], [210, 220], [310, 320]]}, table3)

        self.assertEqual({'keyrows': ['rrrr1'], 'keycols': ['cccc1', 'cccc2', 'cccc3'],
                          'data': [['one', 'two', 'three']]}, table4)

        with self.assertRaises(UserWarning) as e:
            _ = DB.ws('semistrucdata2').ssd()
            self.assertEqual('pylightxl - keyrows != keycols most likely due to missing keyword flag '
                             'keycol IDs: [1], keyrow IDs: []', e)

    def test_new_empty_cell(self):
        self.assertEqual('', DB.ws('empty').index(1, 1))
        DB.set_emptycell(val='NA')
        self.assertEqual('NA', DB.ws('empty').index(1, 1))
        DB.set_emptycell(val=0)
        self.assertEqual(0, DB.ws('empty').index(1, 1))
        # reset it so other tests run correctly
        DB.set_emptycell(val='')

    def test_readingfromIO(self):
        with open('openpyxl.xlsx', 'rb') as f:
            DB = xl.readxl(f)
        

class TestDatabase(TestCase):
    db = xl.Database()

    def test_db_badsheet(self):
        db = xl.Database()
        with self.assertRaises(UserWarning) as e:
            db.ws('not a sheet')
            self.assertEqual('pylightxl - Sheetname (not a sheet) is not in the database', e)

    def test_db_init(self):
        # locally defined to return an empty ws
        db = xl.Database()
        self.assertEqual({}, db._ws)

    def test_db_repr(self):
        self.assertEqual('pylightxl.Database', str(DB))

    def test_db_ws_names(self):
        # locally defined to return an empty list
        db = xl.Database()
        self.assertEqual([], db.ws_names)

    def test_db_add_ws(self):
        db = xl.Database()
        db.add_ws(ws='test1', data={})
        self.assertEqual('pylightxl.Database.Worksheet', str(db.ws(ws='test1')))
        self.assertEqual(['test1'], db.ws_names)
        db.add_ws('test2')
        self.assertEqual(['test1', 'test2'], db.ws_names)

    def test_db_remove_ws(self):
        db = xl.Database()
        db.add_ws('one')
        db.add_ws('two')
        db.add_ws('three')

        db.remove_ws(ws='two')

        self.assertEqual(['one', 'three'], db.ws_names)
        self.assertEqual(False, 'two' in db._ws.keys())
        # remove one thats not in the db
        self.assertEqual(None, db.remove_ws('not real'))

    def test_namedranges(self):
        db = xl.Database()

        # single entry
        db.add_nr(ws='one', name='r1', address='A1')
        self.assertEqual({'r1': 'one!A1'}, db.nr_names)
        # multi entry
        db.add_nr(ws='two', name='r2', address='A2:A3')
        self.assertEqual({'r1': 'one!A1', 'r2': 'two!A2:A3'}, db.nr_names)
        # overwrite by name
        db.add_nr(ws='three', name='r1', address='A3')
        self.assertEqual({'r1': 'three!A3', 'r2': 'two!A2:A3'}, db.nr_names)
        # overwrite by address
        db.add_nr(ws='three', name='r3', address='A3')
        self.assertEqual({'r3': 'three!A3', 'r2': 'two!A2:A3'}, db.nr_names)
        # overwrite by both name and address
        db.add_nr(ws='three', name='r3', address='A4')
        self.assertEqual({'r3': 'three!A4', 'r2': 'two!A2:A3'}, db.nr_names)
        # remove $ references
        db.add_nr(ws='three', name='r3', address='$A$4')
        self.assertEqual({'r3': 'three!A4', 'r2': 'two!A2:A3'}, db.nr_names)

        # remove a nr
        db.remove_nr(name='r3')
        self.assertEqual({'r2': 'two!A2:A3'}, db.nr_names)
        # call a nr that is not in there
        self.assertEqual([[]], db.nr('not real'))

        # check address
        self.assertEqual(['two', 'A2:A3'], db.nr_loc('r2'))

        # update nr
        db.update_nr('r2',10)
        self.assertEqual([[10],[10]], db.nr('r2'))


    def test_namedrange_val(self):
        db = xl.Database()
        db.add_ws('sh1')
        db.ws('sh1').update_address('A1', 11)
        db.ws('sh1').update_address('B1', 12)
        db.ws('sh1').update_address('C2', 23)

        db.add_nr(name='table1', ws='sh1', address='A1')
        db.add_nr(name='table2', ws='sh1', address='A1:C2')

        self.assertEqual([[11]], db.nr(name='table1'))
        self.assertEqual([[11, 12, ''], ['', '', 23]], db.nr(name='table2'))

        db.ws('sh1').update_address('A1', '=11')
        db.ws('sh1').update_address('B1', '=12')
        db.ws('sh1').update_address('C2', '=23')

        self.assertEqual([['=11']], db.nr(name='table1', output='f'))
        self.assertEqual([['=11', '=12', ''], ['', '', '=23']], db.nr(name='table2', output='f'))

    def test_rename_ws(self):
        db = xl.Database()
        db.add_ws('one')
        db.ws('one').update_address('A1', 10)
        db.add_ws('two')
        db.ws('two').update_address('A1', 20)
        db.add_ws('three')
        db.ws('three').update_address('A1', 30)

        # rename to overlapping name should keep the data of the "two", "one" should be removed
        db.rename_ws('one', 'two')
        self.assertEqual(['two', 'three'], db.ws_names)
        self.assertEqual(10, db.ws('two').address('A1'))
        # name a ws thats not in db
        self.assertEqual(None, db.rename_ws('not real', 'new'))
        # rename to new sheet
        db.rename_ws('three', 'four')
        self.assertEqual(['two', 'four'], db.ws_names)


class TestWorksheet(TestCase):

    def test_ws_init(self):
        ws = xl.Worksheet()
        self.assertEqual({}, ws._data)
        self.assertEqual(0, ws.maxrow)
        self.assertEqual(0, ws.maxcol)

    def test_ws_repr(self):
        ws = xl.Worksheet()
        self.assertEqual('pylightxl.Database.Worksheet', str(ws))

    def test_ws_calc_size(self):
        ws = xl.Worksheet()
        self.assertEqual(0, ws.maxrow)
        self.assertEqual(0, ws.maxcol)

        ws.update_address('A1', 11)
        self.assertEqual(1, ws.maxrow)
        self.assertEqual(1, ws.maxcol)

        ws.update_address('A2', 21)
        self.assertEqual(2, ws.maxrow)
        self.assertEqual(1, ws.maxcol)

        ws.update_address('B1', 12)
        self.assertEqual(2, ws.maxrow)
        self.assertEqual(2, ws.maxcol)

        ws.update_address('B2', 22)
        self.assertEqual(2, ws.maxrow)
        self.assertEqual(2, ws.maxcol)

        ws = xl.Worksheet()
        ws.update_address('AA1', 27)
        ws.update_address('AAA1', 703)
        self.assertEqual(1, ws.maxrow)
        self.assertEqual(703, ws.maxcol)

        ws = xl.Worksheet()
        ws.update_address('A1', 1)
        ws.update_address('A1000', 1000)
        ws.update_address('A1048576', 1048576)
        self.assertEqual(1048576, ws.maxrow)
        self.assertEqual(1, ws.maxcol)

        ws = xl.Worksheet()
        ws.update_address('A1', 1)
        ws.update_address('AA1', 27)
        ws.update_address('AAA1', 703)
        ws.update_address('XFD1', 16384)
        ws.update_address('A1048576', 1048576)
        self.assertEqual(1048576, ws.maxrow)
        self.assertEqual(16384, ws.maxcol)

    def test_ws_size(self):
        ws = xl.Worksheet()
        self.assertEqual([0, 0], ws.size)
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        self.assertEqual([2, 1], ws.size)

    def test_ws_address(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        self.assertEqual(11, ws.address(address='A1'))
        self.assertEqual(11, ws.address('$A$1'))
        self.assertEqual(11, ws.address('$A1'))
        self.assertEqual(11, ws.address('A$1'))
        self.assertEqual('', ws.address('A2'))

    def test_ws_index(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        self.assertEqual(11, ws.index(row=1, col=1))
        self.assertEqual('', ws.index(1, 2))

    def test_ws_range(self):
        db = xl.Database()
        db.add_ws('sh1')
        db.ws('sh1').update_address('A1', 11)
        db.ws('sh1').update_address('B1', 12)
        db.ws('sh1').update_address('C2', 23)

        self.assertEqual([[11]], db.ws('sh1').range('A1'))
        self.assertEqual([['']], db.ws('sh1').range('AA1'))
        self.assertEqual([[11, 12]], db.ws('sh1').range('A1:B1'))
        self.assertEqual([[11], ['']], db.ws('sh1').range('A1:A2'))
        self.assertEqual([[11, 12], ['', '']], db.ws('sh1').range('A1:B2'))
        self.assertEqual([[11, 12, ''], ['', '', 23]], db.ws('sh1').range('A1:C2'))
        self.assertEqual([[12, '', ''], ['', 23, ''], ['', '', '']], db.ws('sh1').range('B1:D3'))

        db.ws('sh1').update_address('A1', '=11')
        db.ws('sh1').update_address('B1', '=12')
        db.ws('sh1').update_address('C2', '=23')

        self.assertEqual([['=11']], db.ws('sh1').range('A1', output='f'))
        self.assertEqual([['=11', '=12', ''], ['', '', '=23']],
                         db.ws('sh1').range('A1:C2', output='f'))

    def test_ws_row(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        ws.update_address('B1', 12)
        self.assertEqual([11, 12], ws.row(row=1))
        self.assertEqual([21, ''], ws.row(2))
        self.assertEqual(['', ''], ws.row(3))

        db = xl.Database()
        db.add_ws('sh1')
        db.ws('sh1').update_index(1, 1, '=A1')
        db.ws('sh1').update_index(2, 1, '=A2')
        db.ws('sh1').update_index(2, 2, '=B2')
        self.assertEqual(['=A1', ''], db.ws('sh1').row(1, output='f'))
        self.assertEqual(['=A2', '=B2'], db.ws('sh1').row(2, output='f'))

    def test_ws_col(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        ws.update_address('B1', 12)
        self.assertEqual([11, 21], ws.col(col=1))
        self.assertEqual([12, ''], ws.col(2))
        self.assertEqual(['', ''], ws.col(3))

        db = xl.Database()
        db.add_ws('sh1')
        db.ws('sh1').update_index(1, 1, '=A1')
        db.ws('sh1').update_index(2, 1, '=A2')
        db.ws('sh1').update_index(2, 2, '=B2')
        self.assertEqual(['=A1', '=A2'], db.ws('sh1').col(1, output='f'))
        self.assertEqual(['', '=B2'], db.ws('sh1').col(2, output='f'))

    def test_ws_rows(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        ws.update_address('B1', 12)
        correct_list = [[11, 12], [21, '']]
        for i, row in enumerate(ws.rows):
            self.assertEqual(correct_list[i], row)

    def test_ws_cols(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        ws.update_address('B1', 12)
        correct_list = [[11, 21], [12, '']]
        for i, col in enumerate(ws.cols):
            self.assertEqual(correct_list[i], col)

    def test_ws_keycol(self):
        ws = xl.Worksheet()
        ws.update_address('A1', 11)
        ws.update_address('A2', 21)
        ws.update_address('A3', 11)
        ws.update_address('B1', 11)
        ws.update_address('B2', 22)
        ws.update_address('B3', 32)
        ws.update_address('C1', 13)
        ws.update_address('C2', 23)
        ws.update_address('C3', 33)

        self.assertEqual([11, 21, 11], ws.keycol(key=11))
        self.assertEqual([11, 21, 11], ws.keycol(key=11, keyindex=1))
        self.assertEqual([], ws.keycol(key=11, keyindex=2))
        self.assertEqual([11, 22, 32], ws.keycol(key=32, keyindex=3))

        self.assertEqual([11, 11, 13], ws.keyrow(key=11))
        self.assertEqual([11, 11, 13], ws.keyrow(key=11, keyindex=1))
        self.assertEqual([11, 11, 13], ws.keyrow(key=11, keyindex=2))
        self.assertEqual([21, 22, 23], ws.keyrow(key=22, keyindex=2))
        self.assertEqual([], ws.keyrow(key=22, keyindex=3))

    def test_update_index(self):
        ws = xl.Worksheet()
        ws.update_index(row=4, col=2, val=42)
        self.assertEqual([4, 2], ws.size)
        self.assertEqual(42, ws.index(4, 2))
        self.assertEqual(42, ws.address('B4'))
        self.assertEqual(42, ws.row(4)[1])
        self.assertEqual(42, ws.col(2)[3])
        # update with empty data
        ws.update_index(1, 1, '')
        self.assertEqual('', ws.index(1, 1))
        # update with formula
        ws.update_index(1, 1, '=A2')
        self.assertEqual('=A2', ws.index(1, 1, output='f'))

    def test_update_address(self):
        ws = xl.Worksheet()
        ws.update_address(address='B4', val=42)
        self.assertEqual([4, 2], ws.size)
        self.assertEqual(42, ws.index(4, 2))
        self.assertEqual(42, ws.address('B4'))
        self.assertEqual(42, ws.row(4)[1])
        self.assertEqual(42, ws.col(2)[3])
        # update with empty data
        ws.update_address('A1', '')
        self.assertEqual('', ws.address('A1'))
        # update with formula
        ws.update_address('A1', '=A2')
        self.assertEqual('=A2', ws.address('A1', output='f'))


class TestConversion(TestCase):

    def test_address2index_baddata(self):
        with self.assertRaises(UserWarning) as e:
            xl.utility_address2index(address=1)
            self.assertEqual('pylightxl - Address (1) must be a string.', e)

        with self.assertRaises(UserWarning) as e:
            xl.utility_address2index('')
            self.assertEqual('pylightxl - Address ('') cannot be an empty str.', e)

        with self.assertRaises(UserWarning) as e:
            xl.utility_address2index('1')
            self.assertEqual('pylightxl - Incorrect address (1) entry. Address must be an alphanumeric '
                                'where the starting character(s) are alpha characters a-z', e)

        with self.assertRaises(UserWarning) as e:
            xl.utility_address2index('1A')
            self.assertEqual('pylightxl - Incorrect address (1A) entry. Address must be an alphanumeric '
                                'where the starting character(s) are alpha characters a-z', e)

        with self.assertRaises(UserWarning) as e:
            xl.utility_address2index('AA')
            self.assertEqual('pylightxl - Incorrect address (AA) entry. Address must be an alphanumeric '
                                'where the trailing character(s) are numeric characters 1-9', e)

    def test_address2index(self):
        self.assertEqual([1, 1], xl.utility_address2index('A1'))
        self.assertEqual([1000, 1], xl.utility_address2index('A1000'))
        self.assertEqual([1048576, 1], xl.utility_address2index('A1048576'))

        self.assertEqual([1, 26], xl.utility_address2index('Z1'))
        self.assertEqual([1, 27], xl.utility_address2index('AA1'))
        self.assertEqual([1, 53], xl.utility_address2index('BA1'))
        self.assertEqual([1, 667], xl.utility_address2index('YQ1'))
        self.assertEqual([1, 703], xl.utility_address2index('AAA1'))
        self.assertEqual([1, 728], xl.utility_address2index('AAZ1'))
        self.assertEqual([1, 11496], xl.utility_address2index('PZD1'))
        self.assertEqual([1, 11685], xl.utility_address2index('QGK1'))
        self.assertEqual([1, 16384], xl.utility_address2index('XFD1'))

        self.assertEqual([1048576, 16384], xl.utility_address2index('XFD1048576'))

    def test_index2address_baddata(self):
        with self.assertRaises(UserWarning) as e:
            xl.utility_index2address(row='', col=1)
            self.assertEqual('pylightxl - Incorrect row ('') entry. Row must either be a int or float', e)
        with self.assertRaises(UserWarning) as e:
            xl.utility_index2address(1, '')
            self.assertEqual('pylightxl - Incorrect col ('') entry. Col must either be a int or float', e)
        with self.assertRaises(UserWarning) as e:
            xl.utility_index2address(0, 0)
            self.assertEqual('pylightxl - Row (0) and Col (0) entry cannot be less than 1', e)

    def test_index2address(self):
        self.assertEqual('A1', xl.utility_index2address(1, 1))
        self.assertEqual('A1000', xl.utility_index2address(1000, 1))
        self.assertEqual('A1048576', xl.utility_index2address(1048576, 1))

        self.assertEqual('Z1', xl.utility_index2address(1, 26))
        self.assertEqual('AA1', xl.utility_index2address(1, 27))
        self.assertEqual('BA1', xl.utility_index2address(1, 53))
        self.assertEqual('YQ1', xl.utility_index2address(1, 667))
        self.assertEqual('AAA1', xl.utility_index2address(1, 703))
        self.assertEqual('AAZ1', xl.utility_index2address(1, 728))
        self.assertEqual('PZD1', xl.utility_index2address(1, 11496))
        self.assertEqual('QGK1', xl.utility_index2address(1, 11685))
        self.assertEqual('XFD1', xl.utility_index2address(1, 16384))

        self.assertEqual('XFD1048576', xl.utility_index2address(1048576, 16384))

    def test_col2num(self):
        self.assertEqual(1, xl.utility_columnletter2num('A'))
        self.assertEqual(26, xl.utility_columnletter2num('Z'))
        self.assertEqual(27, xl.utility_columnletter2num('AA'))
        self.assertEqual(53, xl.utility_columnletter2num('BA'))
        self.assertEqual(667, xl.utility_columnletter2num('YQ'))
        self.assertEqual(702, xl.utility_columnletter2num('ZZ'))
        self.assertEqual(703, xl.utility_columnletter2num('AAA'))
        self.assertEqual(728, xl.utility_columnletter2num('AAZ'))
        self.assertEqual(11496, xl.utility_columnletter2num('PZD'))
        self.assertEqual(11685, xl.utility_columnletter2num('QGK'))
        self.assertEqual(16384, xl.utility_columnletter2num('XFD'))

    def test_num2col(self):
        self.assertEqual('A', xl.utility_num2columnletters(1))
        self.assertEqual('Z', xl.utility_num2columnletters(26))
        self.assertEqual('AA', xl.utility_num2columnletters(27))
        self.assertEqual('BA', xl.utility_num2columnletters(53))
        self.assertEqual('YQ', xl.utility_num2columnletters(667))
        self.assertEqual('ZZ', xl.utility_num2columnletters(702))
        self.assertEqual('AAA', xl.utility_num2columnletters(703))
        self.assertEqual('AAZ', xl.utility_num2columnletters(728))
        self.assertEqual('PZD', xl.utility_num2columnletters(11496))
        self.assertEqual('QGK', xl.utility_num2columnletters(11685))
        self.assertEqual('XFD', xl.utility_num2columnletters(16384))
