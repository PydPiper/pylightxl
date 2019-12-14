import zipfile
import re
from functools import reduce
from os.path import isfile


class Database:

    def __init__(self):
        self._worksheet = {}

    def __repr__(self):
        return 'pylightxl.Database'

    def worksheet(self, sheetname):
        """
        Indexes worksheets within the database
        :param str sheetname: worksheet name
        :return: pylightxl.Database.Worksheet class object
        """

        try:
            return self._worksheet[sheetname]
        except KeyError:
            raise ValueError('Error - Sheetname ({}) is not in the database'.format(sheetname))

    @property
    def worksheetnames(self):
        """
        Returns a list of database stored worksheet names
        :return: list of worksheet names
        """

        return list(self._worksheet.keys())

    def add_worksheet(self, sheetname, data):
        """
        Logs worksheet name and its data in the database
        :param str sheetname: worksheet name
        :param data: dictionary of worksheet cell values (ex: {'A1': 10, 'A2': 20})
        :return: None
        """

        self._worksheet.update({sheetname: Worksheet(data)})


class Worksheet:

    def __init__(self, data):
        self._data = data
        self.maxrow = 0
        self.maxcol = 0
        self._calc_size()

    def __repr__(self):
        return 'pylightxl.Database.Worksheet'

    def _calc_size(self):
        """
        Calculates the size of the worksheet row/col. This only occurs on initialization
        :return: None (but this creates instance attributes maxrow/maxcol)
        """

        if self._data != {}:
            list_of_addresses = list(self._data.keys())
            list_of_addresses.sort()
            largest_col_address = list_of_addresses[-1]
            self.maxcol = self.address2index(largest_col_address)[1]

            strVSnum = re.compile(r'[A-Z]+')
            list_of_rows = [int(strVSnum.split(address)[1]) for address in list_of_addresses]
            self.maxrow = max(list_of_rows)
        else:
            self.maxrow = 0
            self.maxcol = 0

    @property
    def size(self):
        """
        Returns the size of the worksheet (row/col)
        :return: list of [maxrow, maxcol]
        """

        return [self.maxrow, self.maxcol]

    def address(self, address):
        """
        Takes an excel address and returns the worksheet stored value
        :param str address: Excel address (ex: "A1")
        :return: cell value
        """

        try:
            rv = self._data[address]
        except KeyError:
            # no data was parsed, return empty cell value
            rv = ""

        return rv

    def index(self, row, col):
        """
        Takes an excel row and col starting at index 1 and returns the worksheet stored value
        :param int row: row index (starting at 1)
        :param int col: col index (start at 1 that corresponds to column "A")
        :return: cell value
        """

        address = self.index2address(row,col)
        try:
            rv = self._data[address]
        except KeyError:
            # no data was parsed, return empty cell value
            rv = ""

        return rv

    def row(self, row):
        """
        Takes a row index input and returns a list of cell data
        :param int row: row index (starting at 1)
        :return: list of cell data
        """

        rv = []

        for c in range(1, self.maxcol + 1):
            val = self.index(row, c)
            rv.append(val)

        return rv

    def col(self, col):
        """
        Takes a col index input and returns a list of cell data
        :param int col: col index (start at 1 that corresponds to column "A")
        :return: list of cell data
        """

        rv = []

        for r in range(1, self.maxrow + 1):
            val = self.index(r, col)
            rv.append(val)

        return rv

    @property
    def rows(self):
        """
        Returns a list of rows that can be iterated through
        :return: list of rows-lists (ex: [[11,12,13],[21,22,23]] for 2 rows with 3 columns of data
        """

        rv = []

        for r in range(1, self.maxrow + 1):
            rv.append(self.row(r))

        return rv

    @property
    def cols(self):
        """
        Returns a list of cols that can be iterated through
        :return: list of cols-lists (ex: [[11,21],[12,22],[13,23]] for 2 rows with 3 columns of data
        """

        rv = []

        for c in range(1, self.maxcol + 1):
            rv.append(self.col(c))

        return rv

    def address2index(self, address):
        """
        Convert excel address to row/col index
        :param str address: Excel address (ex: "A1")
        :return: list of [row, col]
        """
        if type(address) is not str:
            raise ValueError('Error - Address ({}) must be a string.'.format(address))
        if address == '':
            raise ValueError('Error - Address ({}) cannot be an empty str.'.format(address))

        address = address.upper()

        strVSnum = re.compile(r'[A-Z]+')
        try:
            colstr = strVSnum.findall(address)[0]
        except IndexError:
            raise ValueError('Error - Incorrect address ({}) entry. Address must be an alphanumeric '
                             'where the starting character(s) are alpha characters a-z'.format(address))

        if not colstr.isalpha():
            raise ValueError('Error - Incorrect address ({}) entry. Address must be an alphanumeric '
                             'where the starting character(s) are alpha characters a-z'.format(address))

        if len(colstr) == 1:
            col = ord(colstr) - 64
        else:
            col = reduce(lambda x, y: ord(x)-64+ord(y)-64, colstr)

        try:
            row = int(strVSnum.split(address)[1])
        except (IndexError, ValueError):
            raise ValueError('Error - Incorrect address ({}) entry. Address must be an alphanumeric '
                             'where the trailing character(s) are numeric characters 1-9'.format(address))

        return [row, col]

    def index2address(self, row, col):
        """
        Converts index row/col to excel address
        :param int row: row index (starting at 1)
        :param int col: col index (start at 1 that corresponds to column "A")
        :return: str excel address
        """
        if type(row) is not int and type(row) is not float:
            raise ValueError('Error - Incorrect row ({}) entry. Row must either be a int or float'.format(row))
        if type(col) is not int and type(col) is not float:
            raise ValueError('Error - Incorrect col ({}) entry. Col must either be a int or float'.format(col))
        if row <= 0 or col <= 0:
            raise ValueError('Error - Row ({}) and Col ({}) entry cannot be less than 1'.format(row, col))

        colname = chr(col+64)
        return colname + str(row)


def readxl(fn):
    """
    Reads an xlsx or xlsm file and returns a pylightxl database
    :param str fn: Excel file name
    :return: pylightxl.Database class
    """

    # declare a db
    db = Database()

    # test that file entered was a valid excel file
    check_excelfile(fn)

    # zip up the excel file to expose the xml files
    with zipfile.ZipFile(fn, 'r') as f_zip:

        # get custom sheetnames
        with f_zip.open('xl/workbook.xml', 'r') as f:
            sheetnames = get_sheetnames(f)

        # get all of the zip'ed xml sheetnames
        zip_sheetnames = get_zipsheetnames(f_zip)

        # gat common string cell value table
        if 'xl/sharedStrings.xml' in f_zip.NameToInfo.keys():
            with f_zip.open('xl/sharedStrings.xml') as f:
                sharedString = get_sharedStrings(f)
        else:
            sharedString = {}

        # scrape each sheet#.xml file
        for i, zip_sheetname in enumerate(zip_sheetnames):
            with f_zip.open(zip_sheetname, 'r') as f:
                db.add_worksheet(sheetname=sheetnames[i], data=scrape_worksheetxml(f, sharedString))

    return db


def check_excelfile(fn):
    """
    Takes a file-path and raises error if the file is not found/unsupported.
    :param str fn: Excel file path
    :return: None
    """

    if type(fn) is not str:
        raise ValueError('Error - Incorrect file entry ({}).'.format(fn))

    if not isfile(fn):
        raise ValueError('Error - File ({}) does not exit.'.format(fn))

    extension = fn.split('.')[-1]

    if extension not in ['xlsx', 'xlsm']:
        raise ValueError('Error - Incorrect Excel file extension ({}). '
                         'File extension supported: .xlsx .xlsm'.format(extension))


def get_sheetnames(file):
    """
    Takes a file-handle of xl/workbook.xml and returns a list of sheetnames
    :param open-filehanle file: xl/workbook.xml file-handle
    :return: list of sheetnames
    """

    sheetnames = []

    text = file.read().decode()

    tag_sheets = re.compile(r'(?<=<sheets>)(.*)(?=</sheets>)')
    sheet_section = tag_sheets.findall(text)[0]
    # this will find something like:
    # ['sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="sh2" sheetId="2" r:id']

    # split on '/>' to get each <sheet r.../> as a separate list item
    #   last item on list has to be removed because string ends with '/>'
    sheet_lines = sheet_section.split('/>')[:-1]
    for sheet_line in sheet_lines:
        # split sheet line on '"' will result with: ['sheet name=','Sheet1', 'sheetId=', '1', 'r:id=', 'rId1']
        # simply index to 1 to get the sheet name: Sheet1
        sheetnames.append(sheet_line.split('"')[1])

    return sheetnames

def get_zipsheetnames(zipfile):
    """
    Takes a zip-file-handle and returns a list of default xl sheetnames (ie, Sheet1, Sheet2...)
    :param zip-filehandle zipfile: zip file-handle of the excel file
    :return: list of zip xl sheetname paths
    """

    return [name for name in zipfile.NameToInfo.keys() if 'sheet' in name]


def get_sharedStrings(file):
    """
    Takes a file-handle of xl/sharedStrings.xml and returns a dictionary of commonly used strings
    :param open-filehandle file: xl/sharedString.xml file-handle
    :return: dict of commonly used strings
    """

    sharedStrings = {}

    text = file.read().decode()

    tag_t = re.compile(r'<t>(.*?)</t>')
    tag_t_vals = tag_t.findall(text)
    # this will find each string value already separated out as a list

    for i, val in enumerate(tag_t_vals):
        sharedStrings.update({i: val})

    return sharedStrings


def scrape_worksheetxml(file, sharedString=None):
    """
    Takes a file-handle of xl/worksheets/sheet#.xml and returns a dict of cell data
    :param open-filehandle file: xl/worksheets/sheet#.xml file-handle
    :param dict sharedString: shared string dict lookup table from xl/sharedStrings.xml for string only cell values
    :return: yields a dict of cell data {cellAddress: cellVal}
    """

    data = {}

    sharedString = sharedString if sharedString != None else {}

    text = file.read().decode()

    tag_sheetdata = re.compile(r'(?<=<sheetData>)(.*)(?=</sheetData>)')
    sheetdata_section = tag_sheetdata.findall(text)[0]
    tag_cr = re.compile(r'<c r=')
    tag_cr_lines = tag_cr.split(sheetdata_section)[1:]
    # this will find something like:
    # ['"A1"><v>1</v></c>', '"B1"><v>10.1</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.25">',...]
    # the [1:] at the end is to remove the starting split on '<c r=' that would otherwise give a <row r...
    #   text as index 0 (which does not have an address in it) so we simply remove it

    for tag_cr_line in tag_cr_lines:
        # pull out cell address and test if it's a string cell that needs lookup
        re_cell_address = re.compile(r'[^<r c="][^"]+')
        finding_cell_address = re_cell_address.findall(tag_cr_line)
        cell_address = finding_cell_address[0]
        cell_string = True if 't="s"' in tag_cr_line else False

        re_cell_val = re.compile(r'(?<=<v>)(.*)(?=</v>)')
        cell_val = re_cell_val.findall(tag_cr_line)[0]

        if cell_string is True:
            cell_val = sharedString[int(cell_val)]

        data.update({cell_address: cell_val})

    return data





db = readxl('../Book2.xlsx')
pass

