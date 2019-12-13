import zipfile
import re
from functools import reduce

class Database():

    def __init__(self):
        self.worksheet = {}

    def __repr__(self):
        return 'pylightxl.Database'

    @property
    def worksheetnames(self):
        return list(self.worksheet.keys())


class Worksheet():

    def __init__(self):
        self.data = {}

    @property
    def maxrow(self):
        pass

    def _calc_size(self):
        list_of_addresses = list(self.data.keys()).sort()
        largest_col_address = list_of_addresses[-1]
        col = self.address2index(largest_col_address)[1]
        #TODO: finish calc-ing row
        pass

    @property
    def size(self):
        return [self.maxrow, self.maxcol]

    def address(self, range):
        pass

    def index(self, row, col):
        pass

    def address2index(self, address):
        strVSnum = re.compile(r'[A-Z]+')
        colstr = strVSnum.findall(address)[0]
        col = reduce(lambda x,y: ord(x)-64+ord(y)-64, colstr)
        row = int(strVSnum.split(address)[0])
        return [row,col]


def readxl(fn):
    """
    Reads an xlsx or xlsm file and returns a pylightxl database
    :param str fn: Excel file name
    :return: pylightxl database class
    """

    # declare a db
    db = Database()

    # zip up the excel file to expose the xml files
    with zipfile.ZipFile(fn, 'r') as f_zip:

        # get custom sheetnames
        with f_zip.open('xl/workbook.xml', 'r') as f:
            sheetnames = get_sheetnames(f)
            _ = [db.worksheet.update({sheetname: {}}) for sheetname in sheetnames]

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
                db.worksheet[sheetnames[i]].update(scrape_worksheetxml(f, sharedString))

    return db

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

def cellAddress_to_cellRowCol(cellAddress):
    """
    Takes an excel cell address (ie. "A1") and returns a list of the equivalent row/col
    :param str cellAddress: excel cell address (ie. "A1")
    :return: list of [row,col]
    """

    pass




db = readxl('../Book2.xlsx')
pass

