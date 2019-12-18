from unittest import TestCase
from _src.readxl import readxl

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


