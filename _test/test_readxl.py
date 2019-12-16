from unittest import TestCase
from _src.readxl import readxl


DB = readxl('_test/testbook.xlsx')

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
            db = readxl('_test/test_readxl.py')
            self.assertEqual(e, 'Error - Incorrect Excel file extension ({}). '
                                'File extension supported: .xlsx .xlsm'.format('py'))


class test_readxl_integration(TestCase):

    pass
