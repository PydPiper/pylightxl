import re
import sys
# unicode is a python27 object that was merged into str in 3+, for compatibility it is redefined here
if sys.version_info[0] >= 3:
    unicode = str

# future ideas:
# - function to remove empty rows/cols
# - custom row or col key specification that then working with new functions keyrow keycol to give pandas like dataframe dicts
# - matrix function to output 2D data lists


class Database:

    def __init__(self):
        self._ws = {}
        self._sharedStrings = []

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
        """
        Takes a data dict of worksheet cell data (ex: {'A1': 1})

        :param dict data: worksheet cell data (ex: {'A1': 1})
        """
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
                list_of_nums.append(int(''.join(filter(lambda x: unicode(x).isnumeric(), address))))
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

    def keycol(self, key, keyindex=1):
        """
        Takes a column key value (value of any cell within keyindex row) and returns the entire column,
        no match returns an empty list

        :param str/int/float key: any cell value within keyindex row (type sensitive)
        :param int keyindex: option keyrow override. Must be >0 and smaller than worksheet size
        :return list: list of the entire matched key column data (only first match is returned)
        """

        if not keyindex > 0 and not keyindex <= self.size[0]:
            raise ValueError('Error - keyindex ({}) entered must be >0 and <= worksheet size ({}.'.format(keyindex,self.size))

        # find first key match, get its column index and return col list
        for col_i in range(1, self.size[1] + 1):
            if key == self.index(keyindex, col_i):
                return self.col(col_i)
        return []

    def keyrow(self, key, keyindex=1):
        """
        Takes a row key value (value of any cell within keyindex col) and returns the entire row,
        not match returns an empty list

        :param str/int/float key: any cell value within keyindex col (type sensitive)
        :param int keyrow: option keyrow override. Must be >0 and smaller than worksheet size
        :return list: list of the entire matched key row data (only first match is returned)
        """

        if not keyindex > 0 and not keyindex <= self.size[1]:
            raise ValueError('Error - keyindex ({}) entered must be >0 and <= worksheet size ({}.'.format(keyindex,self.size))

        # find first key match, get its column index and return col list
        for row_i in range(1, self.size[1] + 1):
            if key == self.index(row_i, keyindex):
                return self.row(row_i)
        return []


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
    colname = num2columnletters(col)

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


def num2columnletters(num, power=0):
    """
    Takes a column number and converts it to the equivalent excel column letters

    :param int num: column number
    :param int power: internal power multiplier for recursion
    :return str: excel column letters
    """
    if num <= 26:
        return chr(num % 27 + 64)
    elif num > 26**(power+1):
        power += 1
        next_num = num-26**power
        # this will return the higher (right most char) first
        char = num2columnletters(num=next_num, power=power)
        # then call func again on reminder
        remainder = next_num - (int(next_num / (26**power)) * (26 ** power))
        # for roll over case a 26/26 is supposed to be a Z, not a 0 that is not a char
        char_num = int(remainder/(26**(power-1)))+1 if remainder != 0 else 27
        char_next = chr(char_num + 64) if power-1 > 0 else chr(char_num-1 + 64)
        char_all = char + char_next
    else:
        # +1 because this else loop already rolled over the <= 26 value and ex: 27/26 should give you B
        #  but if we did not +1 it would give you an A, however the back end must be conditioned for Z
        #  for a condition when all chars roll over ex: 702-26=676 676/26=26 should still give a Z
        #  but a +1 would throw it over
        char_num = int(num / (26**power)) + 1 if int(num / (26**power)) != 26 else 26
        return chr(char_num + 64)

    return char_all
