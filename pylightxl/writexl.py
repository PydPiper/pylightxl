# standard lib imports
import zipfile
import os
import shutil
import re
from xml.etree import cElementTree as ET
# local lib imports
from .database import index2address



def xml_namespace(file):
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
        # elif event == "end-ns":
        #     ns_map.pop()
        #     return dict(ns_map)
        # elif event == "start":
        #     return dict(ns_map)
    return dict(ns_map)


def writexl(db, path):
    """
    Writes an excel file from pylightxl.Database

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    if not os.path.isfile(path):
        # write to new excel
        new_writer(db, path)
    else:
        # write to existing excel
        alt_writer(db, path)


def alt_writer(db, path):
    """
    Writes to an existing excel file. Only injects cell overwrites or new/removed sheets

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    filename = path.split('/')[-1]
    filename = filename if filename.split('.')[-1] == 'xlsx' else '.'.join(filename.split('.')[:-1] + ['xlsx'])
    temp_folder = '_pylightxl_' + filename


    # have to extract all first to modify
    with zipfile.ZipFile(path, 'r') as f:
        f.extractall(temp_folder)

    text = alt_app_text(db, temp_folder + '/docProps/app.xml')
    with open(temp_folder + '/docProps/app.xml', 'w') as f:
        f.write(text)

    text = new_workbook_text(db)
    with open(temp_folder + '/xl/workbook.xml', 'w') as f:
        f.write(text)

    # rename sheet#.xml to temp to prevent overwriting
    for file in os.listdir(temp_folder + '/xl/worksheets'):
        if '.xml' in file:
            old_name = temp_folder + '/xl/worksheets/' + file
            new_name = temp_folder + '/xl/worksheets/' + 'temp_' + file
            os.rename(old_name, new_name)
    # get filename to xml rId associations
    sheetref = alt_getsheetref(temp_folder)
    existing_sheetnames = [d['name'] for d in sheetref.values()]

    for shID, sheet_name in enumerate(db.ws_names, 1):
        if sheet_name in existing_sheetnames:
            # get the original sheet
            for subdict in sheetref.values():
                if subdict['name'] == sheet_name:
                    fn = 'temp_' + subdict['filename']

            # rewrite the sheet as if it was new
            text = new_worksheet_text(db, sheet_name)
            # feed altered text to new sheet based on db indexing order
            with open(temp_folder + '/xl/worksheets/sheet{}.xml'.format(shID), 'w') as f:
                f.write(text)
            # remove temp xml sheet file
            os.remove(temp_folder + '/xl/worksheets/{}'.format(fn))
        else:
            # this sheet is new, create a new sheet
            text = new_worksheet_text(db, sheet_name)
            with open(temp_folder + '/xl/worksheets/sheet{shID}.xml'.format(shID=shID), 'w') as f:
                f.write(text)

    # this has to come after sheets for db._sharedStrings to be populated
    text = new_workbookrels_text(db)
    with open(temp_folder + '/xl/_rels/workbook.xml.rels', 'w') as f:
        f.write(text)

    if os.path.isfile(temp_folder + '/xl/sharedStrings.xml'):
        # sharedStrings is always recreated from db._sharedStrings since all sheets are rewritten
        os.remove(temp_folder + '/xl/sharedStrings.xml')
    text = new_sharedStrings_text(db)
    with open(temp_folder + '/xl/sharedStrings.xml', 'w') as f:
        f.write(text)

    text = new_content_types_text(db)
    with open(temp_folder + '/[Content_Types].xml', 'w') as f:
        f.write(text)

    # cleanup files that would cause a "repair" workbook
    try:
        shutil.rmtree(temp_folder + '/xl/ctrlProps')
    except FileNotFoundError:
        pass
    try:
        shutil.rmtree(temp_folder + '/xl/drawings')
    except FileNotFoundError:
        pass
    try:
        shutil.rmtree(temp_folder + '/xl/printerSettings')
    except FileNotFoundError:
        pass
    try:
        os.remove(temp_folder + '/xl/vbaProject.bin')
    except FileNotFoundError:
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
    shutil.move(filename, old_dir)
    os.chdir(old_dir)
    # remove temp folder
    shutil.rmtree(temp_folder)


def alt_app_text(db, filepath):
    """
    Takes a docProps/app.xml and returns a db altered text version of the xml

    :param pylightxl.Database db: pylightxl database that contains data to update xml file
    :param str filepath: file path for docProps/app.xml
    :return str: returns the updated xml text
    """

    # extract text from existing app.xml
    ns = xml_namespace(filepath)
    for prefix, uri in ns.items():
        ET.register_namespace(prefix,uri)

    tree = ET.parse(filepath)
    root = tree.getroot()

    # sheet sizes
    tag_i4 = root.findall('./default:HeadingPairs//vt:i4', ns)[0]
    tag_i4.text = str(len(db.ws_names))
    tag_titles_vector = root.findall('./default:TitlesOfParts/vt:vector', ns)[0]
    tag_titles_vector.set('size', str(len(db.ws_names)))
    # sheet names, remove them then add new ones
    for sheet in root.findall('./default:TitlesOfParts//vt:lpstr', ns):
        root.find('./default:TitlesOfParts/vt:vector', ns).remove(sheet)
    for sheet_name in db.ws_names:
        element = ET.Element('vt:lpstr')
        element.text = sheet_name

        root.find('./default:TitlesOfParts/vt:vector', ns).append(element)

    # reset default namespace
    ET.register_namespace('', ns['default'])

    # roll up entire xml file as text
    text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='unicode')

    return text


def alt_getsheetref(temp_folder):
    """
    Takes a file path for the temp pylightxl uncompressed excel xml files and returns the un-altered
    filenames and rIds

    :param str path: file path to pylightxl_temp
    :return dict: dictionary of filenames {rId: {name: '', filename: ''}}
    """

    sheetref = {}

    # -------------------------------------------------------------
    # get worksheet filenames and Ids
    ns = xml_namespace(temp_folder + '/xl/_rels/workbook.xml.rels')
    for prefix, uri in ns.items():
        ET.register_namespace(prefix,uri)

    tree = ET.parse(temp_folder + '/xl/_rels/workbook.xml.rels')
    root = tree.getroot()

    for element in root.findall('./default:Relationship', ns):
        if 'worksheets/sheet' in element.get('Target'):
            Id = element.get('Id')
            filename = element.get('Target').split('/')[1].replace('"', '')
            sheetref.update({Id: {'name': '', 'filename': filename}})

    # -------------------------------------------------------------
    # get custom worksheet names
    ns = xml_namespace(temp_folder + '/xl/workbook.xml')
    for prefix, uri in ns.items():
        ET.register_namespace(prefix,uri)
    tree = ET.parse(temp_folder + '/xl/workbook.xml')
    root = tree.getroot()

    for element in root.findall('./default:sheets/default:sheet', ns):
        Id = 'rId' + element.get('sheetId')
        sheetref[Id]['name'] = element.get('name')

    return sheetref


def new_writer(db, path):
    """
    Writes to a new excel file. The minimum xml parts are zipped together and converted to an .xlsx

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    filename = path.split('/')[-1]
    filename = filename if filename.split('.')[-1] == 'xlsx' else '.'.join(filename.split('.')[:-1] + ['xlsx'])
    path = '/'.join(path.split('/')[:-1])
    path = path + '/' + filename if path else filename

    with zipfile.ZipFile(path, 'w') as zf:
        text_rels = new_rels_text(db)
        zf.writestr('_rels/.rels', text_rels)

        text_app = new_app_text(db)
        zf.writestr('docProps/app.xml', text_app)

        text_core = new_core_text(db)
        zf.writestr('docProps/core.xml', text_core)

        text_workbook = new_workbook_text(db)
        zf.writestr('xl/workbook.xml', text_workbook)

        for shID, sheet_name in enumerate(db.ws_names, 1):
            text_worksheet = new_worksheet_text(db, sheet_name)
            zf.writestr('xl/worksheets/sheet{shID}.xml'.format(shID=shID), text_worksheet)

        if db._sharedStrings:
            text_sharedStrings = new_sharedStrings_text(db)
            zf.writestr('xl/sharedStrings.xml', text_sharedStrings)

        # this has to come after new_worksheet_text for db._sharedStrings to be populated
        text_workbookrels = new_workbookrels_text(db)
        zf.writestr('xl/_rels/workbook.xml.rels', text_workbookrels)

        # this has to come after new_worksheet_text for db._sharedStrings to be populated
        text_content_types = new_content_types_text(db)
        zf.writestr('[Content_Types].xml', text_content_types)


def new_rels_text(db):

    # location: /_rels/.rels
    # inserts: -
    xml_base =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\r\n' \
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\r\n' \
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\r\n' \
                '</Relationships>'

    return xml_base


def new_app_text(db):
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


def new_core_text(db):
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


def new_workbookrels_text(db):
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


def new_workbook_text(db):
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


def new_worksheet_text(db, sheet_name):
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
    # inserts: address, str_option (t="s" for sharedStrings or t="str" for formulas), tag_formula, val
    xml_tag_cr = '<c r="{address}" {str_option}>{tag_formula}<v>{val}</v></c>'

    ws_size = db.ws(sheet_name).size
    if ws_size == [0,0] or ws_size == [1,1]:
        sheet_size_address = 'A1'
    else:
        sheet_size_address = 'A1:' + index2address(ws_size[0],ws_size[1])

    many_tag_row = ''
    for rowID, row in enumerate(db.ws(sheet_name).rows, 1):
        many_tag_cr = ''
        tag_cr = False
        num_of_cr_tags_counter = 0
        for colID, val in enumerate(row, 1):
            address = index2address(rowID, colID)
            str_option = ''
            tag_formula = ''
            try:
                readin_formula = db.ws(sheet_name)._data[index2address(rowID, colID)]['f']
            except KeyError:
                readin_formula = ''

            if val != '':
                if type(val) is str and val[0] != '=':
                    str_option = 't="s"'
                    try:
                        # replace val with its sharedStrings index, note sharedString index does start at 0
                        val = db._sharedStrings.index(val)
                    except ValueError:
                        db._sharedStrings.append(val)
                        val = db._sharedStrings.index(val)

                if readin_formula != '':
                    str_option = 't="str"'
                    tag_formula = '<f>{f}</f>'.format(f=readin_formula)
                    tag_formula = tag_formula.replace('&', '&amp;')
                    val = '"pylightxl - open excel file and save it for formulas to calculate"'

                # let val equation overwrite the readin_formula if it exist (this was a manual input equation)
                if type(val) is str and val[0] == '=':
                    # technically if the result of a formula is a str then str_option should be t="str"
                    #   but this designation is not necessary for excel to open
                    str_option = 't="str"'
                    tag_formula = '<f>{f}</f>'.format(f=val[1:])
                    tag_formula = tag_formula.replace('&', '&amp;')
                    val = '"pylightxl - open excel file and save it for formulas to calculate"'

                tag_cr = True
                num_of_cr_tags_counter += 1
                many_tag_cr += xml_tag_cr.format(address=address, str_option=str_option, tag_formula=tag_formula, val=val)

        if tag_cr:
            many_tag_row += xml_tag_row.format(row_num=rowID, num_of_cr_tags=str(num_of_cr_tags_counter),
                                               many_tag_cr=many_tag_cr)

    # not 100% what uid does, but it is required for excel to open
    rv = xml_base.format(sizeAddress=sheet_size_address, uid='2C7EE24B-C535-494D-AA97-0A61EE84BA40', many_tag_row=many_tag_row)
    return rv


def new_sharedStrings_text(db):
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


def new_content_types_text(db):
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

