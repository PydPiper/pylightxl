import re


class Database:

    def __init__(self):
        self._ws = {}

    def __repr__(self):
        return 'pylightxl.Database'

    def ws(self, sheetname):
        """
        Indexes worksheets within the database
        :param str sheetname: worksheet name
        :return: pylightxl.Database.Worksheet class object
        """

        try:
            return self._ws[sheetname]
        except KeyError:
            raise ValueError('Error - Sheetname ({}) is not in the database'.format(sheetname))

    @property
    def ws_names(self):
        """
        Returns a list of database stored worksheet names
        :return: list of worksheet names
        """

        return list(self._ws.keys())

    def add_ws(self, sheetname, data):
        """
        Logs worksheet name and its data in the database
        :param str sheetname: worksheet name
        :param data: dictionary of worksheet cell values (ex: {'A1': 10, 'A2': 20})
        :return: None
        """

        self._ws.update({sheetname: Worksheet(data)})


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

            list_of_chars = []
            list_of_nums = []
            for address in list_of_addresses:
                list_of_chars.append(''.join(filter(lambda x: x.isalpha(), address)))
                list_of_nums.append(int(''.join(filter(lambda x: x.isnumeric(), address))))
            self.maxrow = int(max(list_of_nums))
            # of all chars are the same length
            list_of_chars.sort(reverse=True)
            # if chars are different length
            list_of_chars.sort(key=len, reverse=True)
            self.maxcol = address2index(list_of_chars[0]+"1")[1]
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

        address = index2address(row,col)
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

        return iter(rv)

    @property
    def cols(self):
        """
        Returns a list of cols that can be iterated through
        :return: list of cols-lists (ex: [[11,21],[12,22],[13,23]] for 2 rows with 3 columns of data
        """

        rv = []

        for c in range(1, self.maxcol + 1):
            rv.append(self.col(c))

        return iter(rv)


def address2index(address):
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

    col = columnletter2num(colstr)

    try:
        row = int(strVSnum.split(address)[1])
    except (IndexError, ValueError):
        raise ValueError('Error - Incorrect address ({}) entry. Address must be an alphanumeric '
                         'where the trailing character(s) are numeric characters 1-9'.format(address))

    return [row, col]


def index2address(row, col):
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

    # values over 26 are outside the A-Z range, reduce them
    val = col % 26 if col % 26 != 0 else 26

    colname = chr(val+64)
    return colname + str(row)


def columnletter2num(text):
    """
    Takes excel column header string and returns the equivalent column count
    :param str text: excel column (ex: 'AAA' will return 703)
    :return: int of column count
    """
    letter_pos = len(text) - 1
    val = 0
    try:
        val = (ord(text[0].upper())-64) * 26 ** letter_pos
        next_val = columnletter2num(text[1:])
        val = val + next_val
    except IndexError:
        return val
    return val
