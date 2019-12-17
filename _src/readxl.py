# standard lib imports
import zipfile
import re
from os.path import isfile
from time import time
# local lib imports
from .database import Database
# c-python functions
from .readxl_scrape import scrape


def readxl(fn, sheetnames=()):
    """
    Reads an xlsx or xlsm file and returns a pylightxl database
    :param str fn: Excel file name
    :param tuple sheetnames: sheetnames to read into the database, if not specified - all sheets are read
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
            sh_names = get_sheetnames(f)

        # get all of the zip'ed xml sheetnames
        zip_sheetnames = get_zipsheetnames(f_zip)

        # remove all names not in entry sheetnames
        if sheetnames:
            temp = []
            for sn in sheetnames:
                try:
                    pop_index = sh_names.index(sn)
                    temp.append(zip_sheetnames.pop(pop_index))
                except ValueError:
                    raise ValueError('Error - Sheetname ({}) is not in the workbook.'.format(sn))
            zip_sheetnames = temp

        # get common string cell value table
        if 'xl/sharedStrings.xml' in f_zip.NameToInfo.keys():
            with f_zip.open('xl/sharedStrings.xml') as f:
                sharedString = get_sharedStrings(f)
        else:
            sharedString = {}

        # scrape each sheet#.xml file
        for i, zip_sheetname in enumerate(zip_sheetnames):
            with f_zip.open(zip_sheetname, 'r') as f:
                db.add_ws(sheetname=sh_names[i], data=scrape(f, sharedString))

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


