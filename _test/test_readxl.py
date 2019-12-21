# standard lib imports
from unittest import TestCase
# local lib imports
from _src.readxl import readxl
from _src.database import Database, Worksheet, address2index, index2address, columnletter2num

try:
    # running from top level
    DB = readxl('_test/testbook.xlsx')
except ValueError:
    # running within _test folder or in debug
    DB = readxl('testbook.xlsx')


class test_readxl_bad_input(TestCase):

    def test_bad_fn_type(self):
        with self.assertRaises(ValueError) as e:
            db = readxl(fn=1)
            self.assertEqual(e,'Error - Incorrect file entry ({}).'.format('1'))

    def test_bad_fn_exist(self):
        with self.assertRaises(ValueError) as e:
            db = readxl('bad')
            self.assertEqual(e, 'Error - File ({}) does not exit.'.format('bad'))

    def test_bad_fn_ext(self):
        with self.assertRaises(ValueError) as e:
            try:
                db = readxl('_test/test_readxl.py')
            except ValueError:
                db = readxl('test_readxl.py')
            self.assertEqual(e, 'Error - Incorrect Excel file extension ({}). '
                                'File extension supported: .xlsx .xlsm'.format('py'))


class test_readxl_integration(TestCase):

    def test_AllSheetsRead(self):
        self.assertEqual(DB.ws_names,['empty','types','scatter','length','sheet_not_to_read'])

    def test_SelectedSheetReading(self):
        try:
            db = readxl('_test/testbook.xlsx',('empty','types'))
        except ValueError:
            db = readxl('testbook.xlsx',('empty','types'))
        self.assertEqual(db.ws_names,['empty','types'])

    def test_commondString(self):
        # all cells that contain strings (without equations are stored in a commondString.xlm)
        self.assertEqual(DB.ws('types').address('A2'),'copy')
        self.assertEqual(DB.ws('types').address('B3'),'ThreeTwo')
        self.assertEqual(DB.ws('types').address('B4'),'copy')

    def test_ws_empty(self):
        # should not contain any cell data, however the user should be able to index to any cell for ""
        self.assertEqual(DB.ws('empty').index(1,1), '')
        self.assertEqual(DB.ws('empty').index(10,10), '')
        self.assertEqual(DB.ws('empty').size, [0,0])
        self.assertEqual(DB.ws('empty').row(1),[])
        self.assertEqual(DB.ws('empty').col(1),[])

    def test_ws_types(self):
        self.assertEqual(DB.ws('types').index(1,1),11)
        self.assertEqual(DB.ws('types').index(2,1),'copy')
        self.assertEqual(DB.ws('types').index(3,1),31)
        self.assertEqual(DB.ws('types').index(4,1),41)
        self.assertEqual(DB.ws('types').index(5,1),'string from A2 copy')
        self.assertEqual(DB.ws('types').index(6,1),'')

        self.assertEqual(DB.ws('types').index(1,2),12.1)
        self.assertEqual(DB.ws('types').index(2,2),'"22"')
        self.assertEqual(DB.ws('types').index(3,2),'ThreeTwo')
        self.assertEqual(DB.ws('types').index(4,2),'copy')
        self.assertEqual(DB.ws('types').index(5,2),'')

        self.assertEqual(DB.ws('types').index(1,3),'')
        self.assertEqual(DB.ws('types').size, [5,2])

        self.assertEqual(DB.ws('types').row(1),[11,12.1])
        self.assertEqual(DB.ws('types').row(2),['copy','"22"'])
        self.assertEqual(DB.ws('types').row(3),[31,'ThreeTwo'])
        self.assertEqual(DB.ws('types').row(4),[41,'copy'])
        self.assertEqual(DB.ws('types').row(5),['string from A2 copy',''])
        self.assertEqual(DB.ws('types').row(6),['',''])

        self.assertEqual(DB.ws('types').col(1),[11,'copy',31,41,'string from A2 copy'])
        self.assertEqual(DB.ws('types').col(2),[12.1,'"22"','ThreeTwo','copy',''])
        self.assertEqual(DB.ws('types').col(3),['','','','',''])

        for i, row in enumerate(DB.ws('types').rows,start=1):
            self.assertEqual(row,DB.ws('types').row(i))
        for i, col in enumerate(DB.ws('types').cols,start=1):
            self.assertEqual(col,DB.ws('types').col(i))

    def test_ws_scatter(self):
        self.assertEqual(DB.ws('scatter').index(1,1),'')
        self.assertEqual(DB.ws('scatter').index(2,2),22)
        self.assertEqual(DB.ws('scatter').index(3,3),33)
        self.assertEqual(DB.ws('scatter').index(3,4),34)
        self.assertEqual(DB.ws('scatter').index(6,6),66)
        self.assertEqual(DB.ws('scatter').index(5,6),'')

        self.assertEqual(DB.ws('scatter').size,[6,6])

    def test_ws_length(self):
        self.assertEqual(DB.ws('length').size,[1048576,16384])


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

        ws._data={'A1': 11}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1)
        self.assertEqual(ws.maxcol, 1)

        ws._data={'A1': 11, 'A2': 21}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 1)

        ws._data={'A1': 11, 'A2': 21, 'B1': 12}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 2)

        ws._data={'A1': 11, 'A2': 21, 'B1': 12, 'B2': 22}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 2)
        self.assertEqual(ws.maxcol, 2)

        ws._data={'A1': 1, 'AA1': 27, 'AAA1': 703}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1)
        self.assertEqual(ws.maxcol, 703)

        ws._data={'A1': 1, 'A1000': 1000, 'A1048576': 1048576}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1048576)
        self.assertEqual(ws.maxcol, 1)

        ws._data={'A1': 1, 'AA1': 27, 'AAA1': 703, 'XFD1': 16384, 'A1048576': 1048576}
        ws._calc_size()
        self.assertEqual(ws.maxrow, 1048576)
        self.assertEqual(ws.maxcol, 16384)

    def test_ws_size(self):
        ws = Worksheet({})
        self.assertEqual(ws.size,[0,0])
        ws._data={'A1': 11, 'A2': 21}
        ws._calc_size()
        self.assertEqual(ws.size, [2,1])

    def test_ws_address(self):
        ws = Worksheet({'A1':1})
        self.assertEqual(ws.address(address='A1'), 1)
        self.assertEqual(ws.address('A2'), '')

    def test_ws_index(self):
        ws = Worksheet({'A1':1})
        self.assertEqual(ws.index(row=1,col=1), 1)
        self.assertEqual(ws.index(1,2), '')

    def test_ws_row(self):
        ws = Worksheet({'A1': 11, 'A2': 21, 'B1': 12})
        self.assertEqual(ws.row(row=1),[11,12])
        self.assertEqual(ws.row(2),[21,''])
        self.assertEqual(ws.row(3),['',''])

    def test_ws_col(self):
        ws = Worksheet({'A1': 11, 'A2': 21, 'B1': 12})
        self.assertEqual(ws.col(col=1),[11,21])
        self.assertEqual(ws.col(2),[12,''])
        self.assertEqual(ws.col(3),['',''])

    def test_ws_rows(self):
        ws = Worksheet({'A1': 11, 'A2': 21, 'B1': 12})
        correct_list = [[11,12],[21,'']]
        for i, row in enumerate(ws.rows):
            self.assertEqual(row,correct_list[i])

    def test_ws_cols(self):
        ws = Worksheet({'A1': 11, 'A2': 21, 'B1': 12})
        correct_list = [[11,21],[12,'']]
        for i, col in enumerate(ws.cols):
            self.assertEqual(col,correct_list[i])


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
        self.assertEqual(address2index('A1'),[1,1])
        self.assertEqual(address2index('A1000'),[1000,1])
        self.assertEqual(address2index('A1048576'),[1048576,1])

        self.assertEqual(address2index('Z1'),[1,26])
        self.assertEqual(address2index('AA1'),[1,27])
        self.assertEqual(address2index('BA1'),[1,53])
        self.assertEqual(address2index('YQ1'),[1,667])
        self.assertEqual(address2index('AAA1'),[1,703])
        self.assertEqual(address2index('QGK1'),[1,11685])
        self.assertEqual(address2index('XFD1'),[1,16384])

        self.assertEqual(address2index('XFD1048576'),[1048576,16384])

    def test_index2address_baddata(self):
        with self.assertRaises(ValueError) as e:
            index2address(row='',col=1)
            self.assertEqual(e, 'Error - Incorrect row ('') entry. Row must either be a int or float')
        with self.assertRaises(ValueError) as e:
            index2address(1,'')
            self.assertEqual(e, 'Error - Incorrect col ('') entry. Col must either be a int or float')
        with self.assertRaises(ValueError) as e:
            index2address(0,0)
            self.assertEqual(e, 'Error - Row (0) and Col (0) entry cannot be less than 1')

    def test_index2address(self):
        self.assertEqual(index2address(1,1),'A1')
        self.assertEqual(index2address(1000,1),'A1000')
        self.assertEqual(index2address(1048576,1),'A1048576')

        self.assertEqual(index2address(1,26),'Z1')
        self.assertEqual(index2address(1,27),'AA1')
        self.assertEqual(index2address(1,53),'BA1')
        self.assertEqual(index2address(1,667),'YQ1')
        self.assertEqual(index2address(1,703),'AAA1')
        self.assertEqual(index2address(1,11685),'QGK1')
        self.assertEqual(index2address(1,16384),'XFD1')

        self.assertEqual(index2address(1048576,16384),'XFD1048576')

    def test_col2num(self):
        self.assertEqual(columnletter2num('A'),1)
        self.assertEqual(columnletter2num('Z'),26)
        self.assertEqual(columnletter2num('AA'),27)
        self.assertEqual(columnletter2num('BA'),53)
        self.assertEqual(columnletter2num('YQ'),667)
        self.assertEqual(columnletter2num('ZZ'),702)
        self.assertEqual(columnletter2num('AAA'),703)
        self.assertEqual(columnletter2num('QGK'),11685)
        self.assertEqual(columnletter2num('XFD'),16384)

