########################################################################################################
# SEC-00: PREFACE
########################################################################################################
"""

Title: pylightxl

Version: 1.44

Source: https://github.com/PydPiper/pylightxl

Documentation: https://pylightxl.readthedocs.io/en/latest/

Description: Pylightxl is a light-weight Microsoft Excel cell value reader/writer. Its strength over
existing libraries comes from the fact that pylightxl has zero non-standard libraries (zero-dependency),
supports python2-3, and its light-weight single file size makes it favorable to copy pylightxl into
your own projects for true zero-dependency. Please see documentation for full list of capabilities and
limitations.

Developers Notes:
    - always write test cases first
    - strive for simple/intuitive API interface
    - write docstrings with type annotations (unfortunately type-hints are not python2 compatible)
    - write documentation as a function is developed

Code Structure:
    - SEC-00: PREFACE
    - SEC-01: IMPORTS
    - SEC-02: READXL FUNCTIONS
    - SEC-03: WRITEXL FUNCTIONS
    - SEC-04: DATABASE FUNCTIONS
    - SEC-05: UTILITY FUNCTIONS

Future Ideas:
    - function to remove empty rows/cols
    - function that output data in pandas like data format (in-case someone needed to convert to pandas)
    - matrix function to output 2D data lists


"""


########################################################################################################
# SEC-01: IMPORTS
########################################################################################################


import zipfile
import re
import os
import sys
import shutil
from xml.etree import cElementTree as ET


########################################################################################################
# SEC-02: PYTHON2 COMPATIBILITY
########################################################################################################


# unicode is a python27 object that was merged into str in 3+, for compatibility it is redefined here
if sys.version_info[0] < 3:
    FileNotFoundError = IOError
    PermissionError = Exception
    WindowsError = Exception
else:
    unicode = str
    WindowsError = Exception


########################################################################################################
# SEC-03: READXL FUNCTIONS
########################################################################################################

def readxl(fn, ws=()):
    """
    Reads an xlsx or xlsm file and returns a pylightxl database

    :param str fn: Excel file name
    :param tuple ws: sheetnames to read into the database, if not specified - all sheets are read
    :return: pylightxl.Database class
    """

    # declare a db
    db = Database()

    # test that file entered was a valid excel file
    if 'pathlib' in str(type(fn)):
        fn = str(fn)

    readxl_check_excelfile(fn)

    # zip up the excel file to expose the xml files
    with zipfile.ZipFile(fn, 'r') as f_zip:

        # get custom sheetnames
        with f_zip.open('xl/workbook.xml', 'r') as f:
            sh_names = readxl_get_sheetnames(f)

        # get all of the zip'ed xml sheetnames, sort in because python27 reads these out of order
        zip_sheetnames = readxl_get_zipsheetnames(f_zip)
        zip_sheetnames.sort()
        # sort again in case there are more than 9 sheets, otherwise sort will be 1,10,2,3,4
        zip_sheetnames.sort(key=len)

        # remove all names not in entry sheetnames
        if ws != ():
            temp = []
            for sn in ws:
                try:
                    pop_index = sh_names.index(sn)
                    temp.append(zip_sheetnames[pop_index])
                except ValueError:
                    raise ValueError('Error - Sheetname ({}) is not in the workbook.'.format(sn))
            zip_sheetnames = temp

        # get common string cell value table
        if 'xl/sharedStrings.xml' in f_zip.NameToInfo.keys():
            with f_zip.open('xl/sharedStrings.xml') as f:
                sharedString = readxl_get_sharedStrings(f, f_zip)
        else:
            sharedString = {}

        # scrape each sheet#.xml file
        if ws == ():
            for i, zip_sheetname in enumerate(zip_sheetnames):
                with f_zip.open(zip_sheetname, 'r') as f:
                    db.add_ws(ws=str(sh_names[i]), data=readxl_scrape(f, sharedString))
        else:
            for sn, zip_sheetname in zip(ws, zip_sheetnames):
                with f_zip.open(zip_sheetname, 'r') as f:
                    db.add_ws(ws=sn, data=readxl_scrape(f, sharedString))

    return db


def readxl_check_excelfile(fn):
    """
    Takes a file-path and raises error if the file is not found/unsupported.

    :param str fn: Excel file path
    :return: None
    """

    if type(fn) is not str:
        raise ValueError('Error - Incorrect file entry ({}).'.format(fn))

    if not os.path.isfile(fn):
        raise ValueError('Error - File ({}) does not exit.'.format(fn))

    extension = fn.split('.')[-1]

    if extension.lower() not in ['xlsx', 'xlsm']:
        raise ValueError('Error - Incorrect Excel file extension ({}). '
                         'File extension supported: .xlsx .xlsm'.format(extension))


def readxl_get_sheetnames(file):
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


def readxl_get_zipsheetnames(zipfile):
    """
    Takes a zip-file-handle and returns a list of default xl sheetnames (ie, Sheet1, Sheet2...)

    :param zip-filehandle zipfile: zip file-handle of the excel file
    :return: list of zip xl sheetname paths
    """

    # rels files will also be created by excel for printer settings, these should not be logged
    return [name for name in zipfile.NameToInfo.keys() if 'sheet' in name and 'rels' not in name]


def readxl_get_sharedStrings(file, f_zip):
    """
    Takes a file-handle of xl/sharedStrings.xml and returns a dictionary of commonly used strings

    :param open-filehandle file: xl/sharedString.xml file-handle
    :return: dict of commonly used strings
    """


    sharedStrings = {}

    # extract text from existing app.xml
    ns = utility_xml_namespace(file)
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)

    try:
        file.seek(0)
        tree = ET.parse(file)
    except:
        # zipfile from python 2.7.18 comes with zipfile 1.6 doesnt come with file.seek method
        # raises UnsupportedOperation error
        with f_zip.open('xl/sharedStrings.xml') as file:
            tree = ET.parse(file)

    root = tree.getroot()
    pass
    for i, tag_si in enumerate(root.findall('./default:si', ns)):
        tag_t = tag_si.findall('./default:r//default:t', ns)
        if tag_t:
            text = ''.join([tag.text for tag in tag_t])
        else:
            text = tag_si.findall('./default:t', ns)[0].text
        sharedStrings.update({i: text})

    return sharedStrings


def readxl_scrape(f, sharedString):
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


########################################################################################################
# SEC-04: WRITEXL FUNCTIONS
########################################################################################################


def writexl(db, fn):
    """
    Writes an excel file from pylightxl.Database

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str fn: file output path
    :return: None
    """

    if not os.path.isfile(fn):
        # write to new excel
        writexl_new_writer(db, fn)
    else:
        # write to existing excel
        writexl_alt_writer(db, fn)


def writexl_alt_writer(db, path):
    """
    Writes to an existing excel file. Only injects cell overwrites or new/removed sheets

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    filename = os.path.split(path)[-1]
    filename = filename if filename.split('.')[-1] == 'xlsx' else '.'.join(filename.split('.')[:-1] + ['xlsx'])
    temp_folder = '_pylightxl_' + filename


    # have to extract all first to modify
    with zipfile.ZipFile(path, 'r') as f:
        f.extractall(temp_folder)

    text = writexl_alt_app_text(db, temp_folder + '/docProps/app.xml')
    with open(temp_folder + '/docProps/app.xml', 'w') as f:
        f.write(text)


    # rename sheet#.xml to temp to prevent overwriting
    for file in os.listdir(temp_folder + '/xl/worksheets'):
        if '.xml' in file:
            old_name = temp_folder + '/xl/worksheets/' + file
            new_name = temp_folder + '/xl/worksheets/' + 'temp_' + file
            os.rename(old_name, new_name)
    # get filename to xml rId associations
    sheetref = writexl_alt_getsheetref(path_wbrels=temp_folder + '/xl/_rels/workbook.xml.rels',
                                       path_wb=temp_folder + '/xl/workbook.xml')
    existing_sheetnames = [d['name'] for d in sheetref.values()]

    text = writexl_new_workbook_text(db)
    with open(temp_folder + '/xl/workbook.xml', 'w') as f:
        f.write(text)

    for shID, sheet_name in enumerate(db.ws_names, 1):
        if sheet_name in existing_sheetnames:
            # get the original sheet
            for subdict in sheetref.values():
                if subdict['name'] == sheet_name:
                    fn = 'temp_' + subdict['filename']

            # rewrite the sheet as if it was new
            text = writexl_new_worksheet_text(db, sheet_name)
            # feed altered text to new sheet based on db indexing order
            with open(temp_folder + '/xl/worksheets/sheet{}.xml'.format(shID), 'w') as f:
                f.write(text)
            # remove temp xml sheet file
            os.remove(temp_folder + '/xl/worksheets/{}'.format(fn))
        else:
            # this sheet is new, create a new sheet
            text = writexl_new_worksheet_text(db, sheet_name)
            with open(temp_folder + '/xl/worksheets/sheet{shID}.xml'.format(shID=shID), 'w') as f:
                f.write(text)

    # this has to come after sheets for db._sharedStrings to be populated
    text = writexl_new_workbookrels_text(db)
    with open(temp_folder + '/xl/_rels/workbook.xml.rels', 'w') as f:
        f.write(text)

    if os.path.isfile(temp_folder + '/xl/sharedStrings.xml'):
        # sharedStrings is always recreated from db._sharedStrings since all sheets are rewritten
        os.remove(temp_folder + '/xl/sharedStrings.xml')
    text = writexl_new_sharedStrings_text(db)
    with open(temp_folder + '/xl/sharedStrings.xml', 'w') as f:
        f.write(text)

    text = writexl_new_content_types_text(db)
    with open(temp_folder + '/[Content_Types].xml', 'w') as f:
        f.write(text)

    # cleanup files that would cause a "repair" workbook
    try:
        shutil.rmtree(temp_folder + '/xl/ctrlProps')
    except (FileNotFoundError, WindowsError):
        pass
    try:
        shutil.rmtree(temp_folder + '/xl/drawings')
    except (FileNotFoundError, WindowsError):
        pass
    try:
        shutil.rmtree(temp_folder + '/xl/printerSettings')
    except (FileNotFoundError, WindowsError):
        pass
    try:
        os.remove(temp_folder + '/xl/vbaProject.bin')
    except (FileNotFoundError, WindowsError):
        pass

    # remove existing file
    try:
        os.remove(path)
    except PermissionError:
        # file is open
        shutil.rmtree(temp_folder)
        raise UserWarning('Error - Cannot write to existing file ({}) that is already open.'.format(filename))



    # log old wd before changing it to temp folder for zipping
    old_dir = os.getcwd()
    # wd must be change to be within the temp folder to get zipfile to prevent the top level temp folder
    #  from being zipped as well
    os.chdir(temp_folder)
    with zipfile.ZipFile(filename, 'w') as f:
        for root, dirs, files in os.walk('.'):
            for file in files:
                # top level "with" statement already creates a excel file that is seen by os.walk
                #  this check skips that empty zip file from being zipped as well
                if file != filename:
                    f.write(os.path.join(root, file))
    # move the zipped up file out of the temp folder
    try:
        shutil.move(filename, old_dir)
    except Exception:
        os.remove(old_dir + '\\' + filename)
        shutil.move(filename, old_dir)
    os.chdir(old_dir)
    # remove temp folder
    shutil.rmtree(temp_folder)


def writexl_alt_app_text(db, filepath):
    """
    Takes a docProps/app.xml filepath and returns the updated xml text version of it.
    Updates:
        - HeadingPairs/vt:variant/vt:i4 "text" after Worksheets
        - TitlesOfParts/vt:vector named filed "size"
        - TitlesOfParts/vt:vector/vt:lpstr

    :param pylightxl.Database db: pylightxl database that contains data to update xml file
    :param str filepath: file path for docProps/app.xml
    :return str: returns the updated xml text
    """

    # extract text from existing app.xml
    ns = utility_xml_namespace(filepath)
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)
    tree = ET.parse(filepath)
    root = tree.getroot()

    # default declarations (worksheets, named ranges)
    old_ws_count = 0
    old_nr_count = 0

    # TODO: named ranges - update vt:vector size=ws_count + nr_count
    # update: number of worksheets and named ranges for the workbook under "HeadingPairs"
    tags_vt = root.findall('./default:HeadingPairs//vt:variant', ns)
    # each tag_vt:variant should only have 1 vt:i4 tag under it (that's the [0] indexing)
    for i_tag_vt, tag_vt in enumerate(tags_vt):
        try:
            if tag_vt[0].text == "Worksheets":
                old_ws_count = int(tags_vt[i_tag_vt + 1][0].text)
                tags_vt[i_tag_vt + 1][0].text = str(len(db.ws_names))
        except IndexError:
            # ill-formatted xml
            raise UserWarning('pylightxl error - Ill formatted xml on docProps/app.xml.\n'
                              'HeadingPairs/vt:vector/vt:variant Worksheets missing vt:variant pair')
        try:
            # TODO: named ranges - count
            if tag_vt[0].text == "Named Ranges":
                old_nr_count = int(tags_vt[i_tag_vt + 1][0].text)
        except IndexError:
            # ill-formatted xml
            raise UserWarning('pylightxl error - Ill formatted xml on docProps/app.xml.\n'
                              'HeadingPairs/vt:vector/vt:variant Named Ranges missing vt:variant pair')

    # update: number of worksheets and named ranges for the workbook under "TitlesOfParts"
    tag_titles_vector = root.findall('./default:TitlesOfParts/vt:vector', ns)[0]
    # TODO: named ranges - count update
    tag_titles_vector.set('size', str(len(db.ws_names) + old_nr_count))

    # update: remove existing worksheet names, preserve named ranges, add new worksheet names
    # TODO: named ranges - vt:lpstr nr names
    for i_tag_vtlpstr, tag_vtlpstr in enumerate(root.findall('./default:TitlesOfParts//vt:lpstr', ns), 1):
        if i_tag_vtlpstr <= old_ws_count:
            root.find('./default:TitlesOfParts/vt:vector', ns).remove(tag_vtlpstr)
    for sheet_name in db.ws_names[::-1]:
        element = ET.Element('vt:lpstr')
        element.text = sheet_name

        root.find('./default:TitlesOfParts/vt:vector', ns).insert(0, element)

    # reset default namespace
    ET.register_namespace('', ns['default'])

    # roll up entire xml file as text
    text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root).decode()

    return text


def writexl_alt_getsheetref(path_wbrels, path_wb):
    """
    Takes a file path for '/xl/_rels/workbook.xml.rels' and '/xl/workbook.xml' files
    and returns a dictionary relationship between sheetID, name and xml worksheet file path
        rId: xml's indexing between files (workbook.xml.rels defined rID to xml worksheet file path)
        sheetId: order of worksheet in the workbook as it appears in excel
        name: worksheet name as it appears in excel (user defined name)
        filename: xml worksheet file path

    :param str path_wbrels: file path to '/xl/_rels/workbook.xml.rels'
    :param str path_wb: file path to '/xl/workbook.xml'
    :return dict: dictionary of filenames {'rId#': {'name': str, 'filename': str, 'sheetId': int}}
    """

    sheetref = {}

    # -------------------------------------------------------------
    # get worksheet filenames and Ids
    ns = utility_xml_namespace(path_wbrels)
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)
    tree = ET.parse(path_wbrels)
    root = tree.getroot()

    for element in root.findall('./default:Relationship', ns):
        if 'worksheets/sheet' in element.get('Target'):
            rId = element.get('Id')
            filename = element.get('Target').split('/')[1].replace('"', '')
            sheetref.update({rId: {'sheetId': None, 'name': None, 'filename': filename}})

    # -------------------------------------------------------------
    # get custom worksheet names
    ns = utility_xml_namespace(path_wb)
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)
    tree = ET.parse(path_wb)
    root = tree.getroot()

    for element in root.findall('./default:sheets/default:sheet', ns):
        rId = element.get('{' + ns['r'] + '}id')
        sheetref[rId]['name'] = element.get('name')
        sheetref[rId]['sheetId'] = int(element.get('sheetId'))

    return sheetref


def writexl_new_writer(db, path):
    """
    Writes to a new excel file. The minimum xml parts are zipped together and converted to an .xlsx

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    filename = os.path.split(path)[-1]
    filename = filename if filename.split('.')[-1] == 'xlsx' else '.'.join(filename.split('.')[:-1] + ['xlsx'])
    path = '/'.join(os.path.split(path)[:-1])
    path = path + '/' + filename if path else filename

    with zipfile.ZipFile(path, 'w') as zf:
        text_rels = writexl_new_rels_text(db)
        zf.writestr('_rels/.rels', text_rels)

        text_app = writexl_new_app_text(db)
        zf.writestr('docProps/app.xml', text_app)

        text_core = writexl_new_core_text(db)
        zf.writestr('docProps/core.xml', text_core)

        text_workbook = writexl_new_workbook_text(db)
        zf.writestr('xl/workbook.xml', text_workbook)

        for shID, sheet_name in enumerate(db.ws_names, 1):
            text_worksheet = writexl_new_worksheet_text(db, sheet_name)
            zf.writestr('xl/worksheets/sheet{shID}.xml'.format(shID=shID), text_worksheet)

        if db._sharedStrings:
            text_sharedStrings = writexl_new_sharedStrings_text(db)
            zf.writestr('xl/sharedStrings.xml', text_sharedStrings)

        # this has to come after new_worksheet_text for db._sharedStrings to be populated
        text_workbookrels = writexl_new_workbookrels_text(db)
        zf.writestr('xl/_rels/workbook.xml.rels', text_workbookrels)

        # this has to come after new_worksheet_text for db._sharedStrings to be populated
        text_content_types = writexl_new_content_types_text(db)
        zf.writestr('[Content_Types].xml', text_content_types)


def writexl_new_rels_text(db):

    # location: /_rels/.rels
    # inserts: -
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\r\n' \
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\r\n' \
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\r\n' \
                '</Relationships>'

    return xml_base


def writexl_new_app_text(db):
    """
    Returns /docProps/app.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /docProps/app.xml text
    """

    # location: /docProps/app.xml
    # inserts: num_sheets, many_tag_vt
    #  note: sheet name order does not matter
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">\r\n' \
                '<Application>Microsoft Excel</Application>\r\n' \
                '<DocSecurity>0</DocSecurity>\r\n' \
                '<ScaleCrop>false</ScaleCrop>\r\n' \
                '<HeadingPairs>\r\n' \
                    '<vt:vector baseType="variant" size="2">\r\n' \
                        '<vt:variant>\r\n' \
                            '<vt:lpstr>Worksheets</vt:lpstr>\r\n' \
                        '</vt:variant>\r\n' \
                        '<vt:variant>\r\n' \
                            '<vt:i4>{num_sheets}</vt:i4>\r\n' \
                        '</vt:variant>\r\n' \
                    '</vt:vector>\r\n' \
               '</HeadingPairs>\r\n' \
               '<TitlesOfParts>\r\n' \
                   '<vt:vector baseType="lpstr" size="{num_sheets}">\r\n' \
                       '{many_tag_vt}\r\n' \
                   '</vt:vector>\r\n' \
               '</TitlesOfParts>\r\n' \
               '<Company></Company>\r\n' \
               '<LinksUpToDate>false</LinksUpToDate>\r\n' \
               '<SharedDoc>false</SharedDoc>\r\n' \
               '<HyperlinksChanged>false</HyperlinksChanged>\r\n' \
               '<AppVersion>16.0300</AppVersion>\r\n' \
               '</Properties>'

    # location: single tag_sheet insert for xml_base
    # inserts: sheet_name
    tag_vt = '<vt:lpstr>{sheet_name}</vt:lpstr>\r\n'

    num_sheets = len(db.ws_names)
    many_tag_vt = ''
    for sheet_name in db.ws_names:
        many_tag_vt += tag_vt.format(sheet_name=sheet_name)
    rv = xml_base.format(num_sheets=num_sheets, many_tag_vt=many_tag_vt)

    return rv


def writexl_new_core_text(db):
    """
    Returns /docProps/core.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /docProps/core.xml text
    """

    # location: /docProps/core.xml
    # inserts: -
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">\r\n' \
                '<dc:creator>pylightxl</dc:creator>\r\n' \
                '<cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>\r\n' \
                '<dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>\r\n' \
                '<dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>\r\n' \
                '</cp:coreProperties>'

    return xml_base


def writexl_new_workbookrels_text(db):
    """
    Returns /xl/_rels/workbook.xml.rels text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /xl/_rels/workbook.xml.rels text
    """

    # location: /xl/_rels/workbook.xml.rels
    # inserts: many_tag_sheets, tag_sharedStrings, tag_calcChain
    #   sheets first for rId# then theme > styles > sharedStrings
    #   note that theme, style, calcChain is not part of the stack. These don't need to be part of the base xml
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                    '{many_tag_sheets}\r\n' \
                    '{tag_sharedStrings}\r\n' \
                '</Relationships>'

    # location: single tag_sheet insert for xml_base
    # inserts: sheet_num
    #  note: rId is not the order of sheets, it just needs to match workbook.xml
    xml_tag_sheet = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'

    # location: sharedStrings insert for xml_base
    # inserts: ID
    xml_tag_sharedStrings = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId{ID}"/>\r\n'

    many_tag_sheets = ''
    for wsID, _ in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(sheet_num=wsID)
    if db._sharedStrings:
        # +1 to increment +1 from the last sheet ID
        tag_sharedStrings = xml_tag_sharedStrings.format(ID=len(db.ws_names)+1)
    else:
        tag_sharedStrings = ''

    rv = xml_base.format(many_tag_sheets=many_tag_sheets,
                         tag_sharedStrings=tag_sharedStrings)
    return rv


def writexl_new_workbook_text(db):
    """
    Returns xl/workbook.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/workbook.xml text
    """

    # location: xl/workbook.xml
    # inserts: many_tag_sheets
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">\r\n' \
                '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>\r\n' \
                '<workbookPr defaultThemeVersion="166925"/>\r\n' \
                    '<sheets>\r\n' \
                        '{many_tag_sheets}\r\n' \
                    '</sheets>\r\n' \
                    '<calcPr calcId="181029"/>\r\n' \
                '</workbook>'

    # location: worksheet tag for xml_base
    # inserts: name, sheet_id, order_id
    #   note id=rId# is referenced by .rels that points to the file locations of each sheet,
    #        while sheetId is sheet order number, name= is the custom name
    xml_tag_sheet = '<sheet name="{sheet_name}" sheetId="{order_id}" r:id="rId{ref_id}"/>\r\n'

    many_tag_sheets = ''
    for shID, sheet_name in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(sheet_name=sheet_name, order_id=shID, ref_id=shID)
    rv = xml_base.format(many_tag_sheets=many_tag_sheets)
    return rv


def writexl_new_worksheet_text(db, sheet_name):
    """
    Returns xl/worksheets/sheet#.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/worksheets/sheet#.xml text
    """

    # dev note: the reason why db._sharedStrings is defined in here is to take advantage of single time
    #  looping through all of the cell data

    # row size and dyDescent are optional values

    # location: xl/worksheets/sheet#.xml
    # inserts: sizeAddress (ex: A1:B5, if empty then A1), many_tag_row
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{uid}">\r\n' \
                    '<dimension ref="{sizeAddress}"/>\r\n' \
                    '<sheetViews>\r\n' \
                        '<sheetView tabSelected="1" workbookViewId="0"/>\r\n' \
                    '</sheetViews>\r\n' \
                    '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>\r\n' \
                    '<sheetData>\r\n' \
                        '{many_tag_row}\r\n' \
                    '</sheetData>\r\n' \
                    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\r\n' \
                '</worksheet>'

    # location: row tag for xml_base
    # inserts: row_num (ex: 1), num_of_cr_tags (ex: 1:5), many_tag_cr
    xml_tag_row = '<row r="{row_num}" x14ac:dyDescent="0.25" spans="1:{num_of_cr_tags}">{many_tag_cr}</row>\r\n'

    # location: c r tag for xml_tag_row
    # inserts: address, str_option (t="s" for sharedStrings), val
    xml_tag_cr = '<c r="{address}" {str_option}><v>{val}</v></c>'

    ws_size = db.ws(sheet_name).size
    if ws_size == [0,0] or ws_size == [1,1]:
        sheet_size_address = 'A1'
    else:
        sheet_size_address = 'A1:' + utility_index2address(ws_size[0],ws_size[1])

    many_tag_row = ''
    for rowID, row in enumerate(db.ws(sheet_name).rows, 1):
        many_tag_cr = ''
        tag_cr = False
        num_of_cr_tags_counter = 0
        for colID, val in enumerate(row, 1):
            address = utility_index2address(rowID, colID)
            str_option = ''
            cell_formula = ''

            # empty cells are not stored in _data
            try:
                cell_formula = db.ws(sheet_name)._data[address]['f']
            except KeyError:
                pass

            # cell contains a formula
            if cell_formula:
                # cells containing formula must not have a type declaration or a <v> tag
                #   to calculate properly when excel is opened
                tag_formula = '<f>{f}</f>'.format(f=cell_formula)
                tag_formula = tag_formula.replace('&', '&amp;')
                tag_cr = True
                num_of_cr_tags_counter += 1
                many_tag_cr += '<c r="{address}">{tag_formula}</c>'.format(address=address,
                                                                           tag_formula=tag_formula)

            # cell value is string
            elif type(val) is str and val != '':
                str_option = 't="s"'
                try:
                    # replace val with its sharedStrings index,
                    #   note sharedString index does start at 0
                    val = db._sharedStrings.index(val)
                except ValueError:
                    db._sharedStrings.append(val.replace('&', '&amp;'))
                    val = db._sharedStrings.index(val.replace('&', '&amp;'))
                tag_cr = True
                num_of_cr_tags_counter += 1
                many_tag_cr += xml_tag_cr.format(address=address, str_option=str_option, val=val)

            # cell does not contain a formula, it is numeric
            elif val != '':
                # val is numeric
                tag_cr = True
                num_of_cr_tags_counter += 1
                many_tag_cr += xml_tag_cr.format(address=address, str_option=str_option, val=val)

        if tag_cr:
            many_tag_row += xml_tag_row.format(row_num=rowID, num_of_cr_tags=str(num_of_cr_tags_counter),
                                               many_tag_cr=many_tag_cr)

    # not 100% what uid does, but it is required for excel to open
    rv = xml_base.format(sizeAddress=sheet_size_address,
                         uid='2C7EE24B-C535-494D-AA97-0A61EE84BA40',
                         many_tag_row=many_tag_row)
    return rv


def writexl_new_sharedStrings_text(db):
    """
    Returns xl/sharedStrings.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/sharedStrings.xml text
    """

    # location: xl/sharedStrings.xml
    # inserts: sharedString_len, many_tag_si
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<sst uniqueCount="{sharedString_len}" count="{sharedString_len}" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\r\n' \
                    '{many_tag_si}\r\n' \
                '</sst>'

    # location: si tag for xml_base
    # inserts: space_preserve (xml:space="preserve"), val
    #   note leading and trailing spaces requires preserve tag: <t xml:space="preserve"> leadingspace</t>
    xml_tag_si = '<si><t {space_preserve}>{val}</t></si>\r\n'

    sharedString_len = len(db._sharedStrings)

    many_tag_si = ''
    for val in db._sharedStrings:
        if val[0] == ' ' or val[-1] == ' ':
            space_preserve = 'xml:space="preserve"'
        else:
            space_preserve = ''
        many_tag_si += xml_tag_si.format(space_preserve=space_preserve, val=val)

    rv = xml_base.format(sharedString_len=sharedString_len, many_tag_si=many_tag_si)
    return rv


def writexl_new_content_types_text(db):
    """
    Returns [Content_Types].xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: [Content_Types].xml text
    """

    # location: [Content_Types].xml
    # inserts: many_tag_sheets, tag_sharedStrings
    #  note calcChain is part of this but it is not necessary for excel to open
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\r\n' \
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\r\n' \
                    '<Default Extension="xml" ContentType="application/xml"/>\r\n' \
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\r\n' \
                    '{many_tag_sheets}\r\n' \
                    '{tag_sharedStrings}\r\n' \
                    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\r\n' \
                    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\r\n' \
                '</Types>'


    xml_tag_sheet = '<Override PartName="/xl/worksheets/sheet{sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\r\n'

    xml_tag_sharedStrings = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\r\n'

    many_tag_sheets = ''
    for sheet_id, _ in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(sheet_id=sheet_id)

    if db._sharedStrings:
        tag_sharedStrings = xml_tag_sharedStrings
    else:
        tag_sharedStrings = ''

    rv = xml_base.format(many_tag_sheets=many_tag_sheets,
                         tag_sharedStrings=tag_sharedStrings)

    return rv


########################################################################################################
# SEC-05: DATABASE FUNCTIONS
########################################################################################################


class Database:

    def __init__(self):
        self._ws = {}
        self._sharedStrings = []
        # list to preserve insertion order <3.6 and easier to reorder for users than keys of dict
        self._ws_names = []

    def __repr__(self):
        return 'pylightxl.Database'

    def ws(self, ws):
        """
        Indexes worksheets within the database

        :param str ws: worksheet name
        :return: pylightxl.Database.Worksheet class object
        """

        try:
            return self._ws[ws]
        except KeyError:
            raise ValueError('Error - Sheetname ({}) is not in the database'.format(ws))

    @property
    def ws_names(self):
        """
        Returns a list of database stored worksheet names

        :return: list of worksheet names
        """

        return self._ws_names

    def add_ws(self, ws, data=None):
        """
        Logs worksheet name and its data in the database

        :param str ws: worksheet name
        :param data: dictionary of worksheet cell values (ex: {'A1': {'v':10,'f':'','s':''}, 'A2': {'v':20,'f':'','s':''}})
        :return: None
        """

        if data is None:
            data = {'A1': {'v': '', 'f': '', 's': ''}}
        self._ws.update({ws: Worksheet(data)})
        self._ws_names.append(ws)

    def set_emptycell(self, val):
        """
        Custom definition for how pylightxl returns an empty cell

        :param val: (default='') empty cell value
        :return: None
        """

        for ws in self.ws_names:
            self.ws(ws).set_emptycell(val)

#TODO: add rename sheet
#TODO: add remove sheet
class Worksheet():

    def __init__(self, data):
        """
        Takes a data dict of worksheet cell data (ex: {'A1': 1})

        :param dict data: worksheet cell data (ex: {'A1': 1})
        """
        self._data = data
        self.maxrow = 0
        self.maxcol = 0
        self._calc_size()
        self.emptycell = ''

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
            # if all chars are the same length
            list_of_chars.sort(reverse=True)
            # if chars are different length
            list_of_chars.sort(key=len, reverse=True)
            self.maxcol = utility_address2index(list_of_chars[0]+"1")[1]
        else:
            self.maxrow = 0
            self.maxcol = 0

    def set_emptycell(self, val):
        """
        Custom definition for how pylightxl returns an empty cell

        :param val: (default='') empty cell value
        :return: None
        """

        self.emptycell = val

    @property
    def size(self):
        """
        Returns the size of the worksheet (row/col)

        :return: list of [maxrow, maxcol]
        """

        return [self.maxrow, self.maxcol]

    def address(self, address, formula=False):
        """
        Takes an excel address and returns the worksheet stored value

        :param str address: Excel address (ex: "A1")
        :param bool formula: flag to return the formula of this cell
        :return: cell value
        """

        try:
            if not formula:
                rv = self._data[address]['v']
            else:
                rv = '=' + self._data[address]['f']
        except KeyError:
            # no data was parsed, return empty cell value
            rv = self.emptycell

        return rv

    def index(self, row, col, formula=False):
        """
        Takes an excel row and col starting at index 1 and returns the worksheet stored value

        :param int row: row index (starting at 1)
        :param int col: col index (start at 1 that corresponds to column "A")
        :param bool formula: flag to return the formula of this cell
        :return: cell value
        """

        address = utility_index2address(row, col)
        try:
            if not formula:
                rv = self._data[address]['v']
            else:
                rv = '=' + self._data[address]['f']
        except KeyError:
            # no data was parsed, return empty cell value
            rv = self.emptycell

        return rv

    def update_index(self, row, col, val):
        """
        Update worksheet data via index

        :param int row: row index
        :param int col: column index
        :param int/float/str val: cell value; equations are strings and must begin with "="
        :return: None
        """
        address = utility_index2address(row, col)
        self.maxcol = col if col > self.maxcol else self.maxcol
        self.maxrow = row if row > self.maxrow else self.maxrow
        # log formulas under formulas and trim off the '='
        if type(val) is str and val[0] == '=':
            # overwrite existing cell val to be empty (it will calc when excel is opened)
            self._data.update({address: {'v': '', 'f': val[1:], 's': ''}})
        else:
            self._data.update({address: {'v': val, 'f': '', 's': ''}})

    def update_address(self, address, val):
        """
        Update worksheet data via address

        :param str address: excel address (ex: "A1")
        :param int/float/str val: cell value; equations are strings and must begin with "="
        :return: None
        """
        row, col = utility_address2index(address)
        self.maxcol = col if col > self.maxcol else self.maxcol
        self.maxrow = row if row > self.maxrow else self.maxrow
        # log formulas under formulas and trim off the '='
        if type(val) is str and val[0] == '=':
            # overwrite existing cell val to be empty (it will calc when excel is opened)
            self._data.update({address: {'v': '', 'f': val[1:], 's': ''}})
        else:
            self._data.update({address: {'v': val, 'f': '', 's': ''}})

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
        no match returns an empty list

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

    def ssd(self, keyrows='KEYROWS', keycols='KEYCOLS'):
        """
        Runs through the worksheet and looks for "KEYROWS" and "KEYCOLS" flags in each cell to identify
        the start of a semi-structured data. A data table is read until an empty header is
        found by row or column. The search supports multiple tables.

        :param str keyrows: (default='KEYROWS') a flag to indicate the start of keyrow's
                            cells below are read until an empty cell is reached
        :param str keycols: (default='KEYCOLS') a flag to indicate the start of keycol's
                            cells to the right are read until an empty cell is reached
        :return list: list of data dict in the form of [{'keyrows': [], 'keycols': [], 'data': [[], ...]}, {...},]
        """

        # find the index of keyrow(s) and keycol(s) plural if there are multiple datasets - this is a fast loop downselect
        kr_colIDs = [colID for colID, col in enumerate(self.cols, 1) if keyrows in col or keyrows + keycols in col or keycols + keyrows in col]
        kc_rowIDs = [rowID for rowID, row in enumerate(self.rows, 1) if keycols in row or keyrows + keycols in row or keycols + keyrows in row]

        # look for duplicate key flags within rows/cols - this is a slower loop
        temp = []
        for row_id in range(1, self.maxrow):
            for col_id in kr_colIDs:
                cell = self.index(row_id, col_id)
                if cell != '' and type(cell) is str and keyrows in cell:
                    temp.append([row_id, col_id])
        kr_indexIDs = temp

        # for col_id in kr_colIDs:
        #     for row_id, cell in enumerate(self.col(col_id), 1):
        #         if cell != '' and type(cell) is str and keyrows in cell:
        #             temp.append([row_id, col_id])

        temp = []
        for row_id in kc_rowIDs:
            for col_id, cell in enumerate(self.row(row_id), 1):
                if cell != '' and type(cell) is str and keycols in cell:
                    temp.append([row_id, col_id])
        kc_indexIDs = temp

        if len(kr_indexIDs) != len(kc_indexIDs):
            raise ValueError('Error - keyrows != keycols most likely due to missing keyword '
                             'flag keyrow IDs: {}, keycol IDs: {}'.format(kr_indexIDs, kc_indexIDs))

        # datas structure: [{'keycols': ..., 'keyrows': ..., 'data'},...]
        datas = []
        dataset_i = 0
        for kr_indexID, kc_indexID in zip(kr_indexIDs, kc_indexIDs):

            r, c = 0, 1

            datas.append({'keyrows': [], 'keycols': [], 'data': []})

            # pull the column for keycol_ID
            kr_header = self.col(kr_indexID[c])[kr_indexID[r]:]
            # find the end for column header (by looking for empty cell)
            try:
                end_col_index = kr_header.index('')
            except ValueError:
                end_col_index = self.maxrow - kr_indexID[r]
            kr_end = end_col_index

            # pull the row for keyrow_ID
            kc_header = self.row(kc_indexID[r])[kc_indexID[c]:]
            # find the end for column header (by looking for empty cell)
            try:
                end_row_index = kc_header.index('')
            except ValueError:
                end_row_index = self.maxrow - kc_indexID[c]
            kc_end = end_row_index

            # truncate headers down
            datas[dataset_i]['keyrows'] = kr_header[:kr_end]
            datas[dataset_i]['keycols'] = kc_header[:kc_end]

            for row_i in range(kr_indexID[r] + 1, kr_indexID[r] + kr_end + 1):
                datas[dataset_i]['data'].append(self.row(row_i)[kc_indexID[c]:kc_indexID[c] + kc_end])
            dataset_i += 1

        return datas


########################################################################################################
# SEC-06: UTILITY FUNCTIONS
########################################################################################################


def utility_address2index(address):
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

    col = utility_columnletter2num(colstr)

    try:
        row = int(strVSnum.split(address)[1])
    except (IndexError, ValueError):
        raise ValueError('Error - Incorrect address ({}) entry. Address must be an alphanumeric '
                         'where the trailing character(s) are numeric characters 1-9'.format(address))

    return [row, col]


def utility_index2address(row, col):
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
    colname = utility_num2columnletters(col)

    return colname + str(row)


def utility_columnletter2num(text):
    """
    Takes excel column header string and returns the equivalent column count

    :param str text: excel column (ex: 'AAA' will return 703)
    :return: int of column count
    """
    letter_pos = len(text) - 1
    val = 0
    try:
        val = (ord(text[0].upper())-64) * 26 ** letter_pos
        next_val = utility_columnletter2num(text[1:])
        val = val + next_val
    except IndexError:
        return val
    return val


def utility_num2columnletters(num):
    """
    Takes a column number and converts it to the equivalent excel column letters

    :param int num: column number
    :return str: excel column letters
    """

    def pre_num2alpha(num):
        if num % 26 != 0:
            num = [num // 26, num % 26]
        else:
            num = [(num - 1) // 26, 26]

        if num[0] > 26:
            num = pre_num2alpha(num[0]) + [num[1]]
        else:
            num = list(filter(lambda x: False if x == 0 else True, num))

        return num

    return "".join(list(map(lambda x: chr(x + 64), pre_num2alpha(num))))


def utility_xml_namespace(file):
    """
    Takes an xml file and returns the root namespace as a dict

    :param str file: xml file path
    :return dict: dictionary of root namespace
    """

    events = "start", "start-ns", "end-ns"

    ns_map = []

    for event, elem in ET.iterparse(file, events):
        if event == "start-ns":
            elem = ('default', elem[1]) if elem[0] == '' else elem
            ns_map.append(elem)
    return dict(ns_map)
