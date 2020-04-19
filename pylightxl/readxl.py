# standard lib imports
import zipfile
import re
from os.path import isfile
#from time import time
# local lib imports
from .database import Database


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
    if 'pathlib' in str(type(fn)):
        fn = str(fn)

    check_excelfile(fn)

    # zip up the excel file to expose the xml files
    with zipfile.ZipFile(fn, 'r') as f_zip:

        # get custom sheetnames
        with f_zip.open('xl/workbook.xml', 'r') as f:
            sh_names = get_sheetnames(f)

        # get all of the zip'ed xml sheetnames, sort in because python27 reads these out of order
        zip_sheetnames = get_zipsheetnames(f_zip)
        zip_sheetnames.sort()
        # sort again in case there are more than 9 sheets, otherwise sort will be 1,10,2,3,4
        zip_sheetnames.sort(key=len)

        # remove all names not in entry sheetnames
        if sheetnames != ():
            temp = []
            for sn in sheetnames:
                try:
                    pop_index = sh_names.index(sn)
                    temp.append(zip_sheetnames[pop_index])
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
        if sheetnames == ():
            for i, zip_sheetname in enumerate(zip_sheetnames):
                with f_zip.open(zip_sheetname, 'r') as f:
                    db.add_ws(sheetname=str(sh_names[i]), data=scrape(f, sharedString))
        else:
            for sn, zip_sheetname in zip(sheetnames, zip_sheetnames):
                with f_zip.open(zip_sheetname, 'r') as f:
                    db.add_ws(sheetname=sn, data=scrape(f, sharedString))

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

    if extension.lower() not in ['xlsx', 'xlsm']:
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

    tag_sheets = re.compile(r'(?<=<sheets>)([\s\S]*)(?=</sheets>)')
    sheet_section = tag_sheets.findall(text)[0].strip()
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

    # rels files will also be created by excel for printer settings, these should not be logged
    return [name for name in zipfile.NameToInfo.keys() if 'sheet' in name and 'rels' not in name]


def get_sharedStrings(file):
    """
    Takes a file-handle of xl/sharedStrings.xml and returns a dictionary of commonly used strings

    :param open-filehandle file: xl/sharedString.xml file-handle
    :return: dict of commonly used strings
    """

    sharedStrings = {}

    text = file.read().decode()
    # remove next lines that mess up re findall
    text = text.replace('\r','')
    text = text.replace('\n','')

    # allowed to search <t... because of <t xml:space="preserve"> call for keeping white spaces
    tag_t = re.compile(r'<t(.*?)</t>')
    tag_t_vals = tag_t.findall(text)
    # this will find each string value already separated out as a list

    for i, val in enumerate(tag_t_vals):
        # remove extras from re finding
        val = val[1:] if 'xml:space="preserve">' not in val else val[22:]
        sharedStrings.update({i: val})

    return sharedStrings


def scrape(f, sharedString):
    """
    Takes a file-handle of xl/worksheets/sheet#.xml and returns a dict of cell data

    :param open-filehandle file: xl/worksheets/sheet#.xml file-handle
    :param dict sharedString: shared string dict lookup table from xl/sharedStrings.xml for string only cell values
    :return: yields a dict of cell data {cellAddress: cellVal}
    """


    data = {}

    sample_size = 10000

    re_cr_tag = re.compile(r'(?<=<c r=)(.+?)(?=</c>)')
    re_cell_val = re.compile(r'(?<=<v>)(.*)(?=</v>)')
    re_cell_formula = re.compile(r'(?<=<f>)(.*)(?=</f>)')

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

        # capture further breakdown of xml where <c r .... /> is used and re_cr_tag doesnt split it
        # re.compile(r'(?<=<c r=)(.+?)(?=</c>|/>)') was removed since it was prematurely splitting c r tags
        # when a formula is closed by a /> as well
        match_splits = []

        while True:
            if match or len(match_splits) != 0:
                if len(match_splits) == 0:
                    first_match = match.pop(0)
                else:
                    first_match = match_splits.pop(0)
                    if '<c r' in first_match:
                        temp = first_match.split('<c r')[0]
                        match_splits += first_match.split('<c r')[1:]
                        first_match = temp
                if '<c r=' in first_match:
                    match_splits = first_match.split('<c r=')
                    continue
                cell_address = str(first_match.split('"')[1])
                is_commonString = True if 't="s"' in first_match else False
                # bool "FALSE" "TRUE" is not logged as a commonString in xml, 0 == FALSE, 1 == TRUE
                is_bool = True if 't="b"' in first_match else False
                # 't="e"' is for error cells "#N/A"
                is_string = True if 't="str"' in first_match or 't="e"' in first_match else False

                try:
                    cell_val = str(re_cell_val.findall(first_match)[0])
                except IndexError:
                    # current cell doesn't have a value
                    cell_val = ''
                    is_string = True

                try:
                    cell_formula = str(re_cell_formula.findall(first_match)[0])
                except IndexError:
                    # current tag does not have a formula
                    cell_formula = ''

                if is_commonString:
                    cell_val = str(sharedString[int(cell_val)])
                elif is_bool:
                    cell_val = 'True' if cell_val == '1' else 'False'
                elif not is_commonString and not is_string:
                    if cell_val.isdigit():
                        cell_val = int(cell_val)
                    else:
                        cell_val = float(cell_val)

                data.update({cell_address: {'v': cell_val, 'f': cell_formula, 's': ''}})
            else:
                # only carry forward the reminder unmatched text

                text_buff = re_cr_tag.split(text_buff)[-1]

                next_buff = f.read(sample_size).decode()
                text_buff += next_buff

                break

        if not next_buff:
            break

    return data
