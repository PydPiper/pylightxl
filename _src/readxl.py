# standard lib imports
import zipfile
import re
from os.path import isfile
# local lib imports
from .database import Database


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
                db.add_worksheet(sheetname=sheetnames[i], data=scrape(f, sharedString))

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


def scrape(f, sharedString=None):

    data = {}
    sharedString = sharedString if sharedString != None else {}


    sample_size = 1000

    re_cr_tag = re.compile(r'(?<=<c r=)(.+?)(?=</c>)')
    re_cell_val = re.compile(r'(?<=<v>)(.*)(?=</v>)')

    # read and dump data till "sheetData" is reached
    while True:
        text_buff = f.read(sample_size).decode()

        # if sample reading catches "sheetData" entirely
        if 'sheetData' in text_buff:
            break
        else:
            # it is possible to slice through "sheetData" during sampling but 2x slices cannot miss
            #   "sheetData" b/c len("sheetData")=9 char which is way less than 2x sample_size
            text_buff += f.read(sample_size).decode()
            if 'sheetData' in text_buff:
                break
            # if "sheetData" was not found, dump text_buff from memory

    # "sheetData" reach, log address/val
    while True:

        match = re_cr_tag.findall(text_buff)

        while True:
            if match:
                first_match = match.pop(0)
                cell_address = first_match.split('"')[1]
                val_common_str = True if 't="s"' in first_match else False

                cell_val = re_cell_val.findall(first_match)[0]
                if val_common_str:
                    cell_val = sharedString[int(cell_val)]

                data.update({cell_address: cell_val})
            else:
                # only carry forward the reminder unmatched text
                text_buff = re_cr_tag.split(text_buff)[-1]
                next_buff = f.read(sample_size).decode()
                text_buff += next_buff
                break

        if not next_buff:
            break

    return data
