# standard lib imports
import zipfile
from os.path import isfile
from os import rename
# local lib imports
from .database import index2address


def writexl(db, path):
    """
    Writes an excel file from pylightxl.Database

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    if not isfile(path):
        # write to new excel
        new_writer(db, path)
    else:
        # write to existing excel
        alter_writer(db, path)


def alter_writer(db, path):
    """
    Writes to an existing excel file. Only injects cell overwrites or new/removed sheets

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

    # TODO: finish alter excel file
    pass


def new_writer(db, path):
    """
    Writes to a new excel file. The minimum xml parts are zipped together and converted to an .xlsx

    :param pylightxl.Database db: database contains sheetnames, and their data
    :param str path: file output path
    :return: None
    """

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
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''.replace('\n', '\r\n')

    return xml_base


def new_app_text(db):
    """
    Returns /docProps/app.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /docProps/app.xml text
    """

    # location: /docProps/app.xml
    # inserts: num_sheets, many_tag_vt
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector baseType="variant" size="2">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>{num_sheets}</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector baseType="lpstr" size="{num_sheets}">
{many_tag_vt}
</vt:vector>
</TitlesOfParts>
<Company></Company>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>'''.replace('\n', '\r\n')

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
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
<dc:creator>pylightxl</dc:creator>
<cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>
</cp:coreProperties>'''.replace('\n', '\r\n')

    return xml_base


def new_workbookrels_text(db):
    """
    Returns /xl/_rels/workbook.xml.rels text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /xl/_rels/workbook.xml.rels text
    """

    # location: /xl/_rels/workbook.xml.rels
    # inserts: many_tag_sheets, tag_sharedStrings, tag_calcChain
    #   sheets first for rId# then theme > styles > sharedStrings > calcChain
    #   note that theme and style is not part of the stack. These don't need to be part of the base xml
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{many_tag_sheets}
{tag_sharedStrings}
{tag_calcChain}
</Relationships>'''.replace('\n', '\r\n')

    # location: single tag_sheet insert for xml_base
    # inserts: sheet_num
    xml_tag_sheet = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'

    # location: sharedStrings insert for xml_base
    # inserts: ID
    xml_tag_sharedStrings = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId{ID}"/>\r\n'

    # location: calcChain insert for xml_base
    # inserts: ID
    #   this will be un-used till cell formulas are supported
    # TODO: add support for formulas at a later time (after writer new/existing are working)
    xml_tag_calcChain = '<Relationship Target="calcChain.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Id="rId{ID}"/>\r\n'

    many_tag_sheets = ''
    for wsID, _ in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(sheet_num=wsID)
    if db._sharedStrings:
        # +1 to increment +1 from the last sheet ID
        tag_sharedStrings = xml_tag_sharedStrings.format(ID=len(db.ws_names)+1)
    else:
        tag_sharedStrings = ''

    rv = xml_base.format(many_tag_sheets=many_tag_sheets,
                         tag_sharedStrings=tag_sharedStrings,
                         tag_calcChain='')
    return rv


def new_workbook_text(db):
    """
    Returns xl/workbook.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/workbook.xml text
    """

    # location: xl/workbook.xml
    # inserts: many_tag_sheets
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>
<workbookPr defaultThemeVersion="166925"/>
<sheets>
{many_tag_sheets}
</sheets>
<calcPr calcId="181029"/>
</workbook>'''.replace('\n', '\r\n')

    # location: worksheet tag for xml_base
    # inserts: name, sheet_id, order_id
    #   note id=rId# is the worksheet order, while sheetId is the worksheet true ID, name= is the custom name
    xml_tag_sheet = '<sheet name="{sheet_name}" sheetId="{sheet_id}" r:id="rId{order_id}"/>\r\n'

    many_tag_sheets = ''
    for shID, sheet_name in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(sheet_name=sheet_name, sheet_id=shID, order_id=shID)
    rv = xml_base.format(many_tag_sheets=many_tag_sheets)
    return rv


def new_worksheet_text(db, sheet_name):
    """
    Returns xl/worksheets/sheet#.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/worksheets/sheet#.xml text
    """

    # location: xl/worksheets/sheet#.xml
    # inserts: sizeAddress (ex: A1:B5, if empty then A1), many_tag_row
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{uid}">
<dimension ref="{sizeAddress}"/>
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0"/>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
<sheetData>
{many_tag_row}
</sheetData>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>'''.replace('\n', '\r\n')

    # location: row tag for xml_base
    # inserts: row_num (ex: 1), row_span (ex: 1:5), many_tag_cr
    xml_tag_row = '<row r="{row_num}" x14ac:dyDescent="0.25" spans="{row_span}">{many_tag_cr}</row>\r\n'

    # location: c r tag for xml_tag_row
    # inserts: address, str_option (t="s" for sharedStrings or t="str" for formulas), val
    #   currently formulas are unsupported
    # TODO: add support for formulas at a later time (after writer new/existing are working)
    xml_tag_cr = '<c r="{address}" {str_option}><v>{val}</v></c>'

    ws_size = db.ws(sheet_name).size
    if ws_size == [0,0] or ws_size == [1,1]:
        sheet_size_address = 'A1'
    else:
        sheet_size_address = 'A1:' + index2address(ws_size[0],ws_size[1])

    many_tag_row = ''
    for rowID, row in enumerate(db.ws(sheet_name).rows, 1):
        many_tag_cr = ''
        tag_cr = False
        for colID, val in enumerate(row, 1):
            address = index2address(rowID, colID)
            # TODO: update str_option when formulas are supported
            if type(val) is str and val != '':
                str_option = 's'
                try:
                    # replace val with its sharedStrings index, +1 since python starts at 0
                    val = db._sharedStrings.index(val) + 1
                except ValueError:
                    db._sharedStrings.append(val)
                    # +1 since python starts at 0
                    val = db._sharedStrings.index(val) + 1
            else:
                str_option = ''
            if val != '':
                tag_cr = True
                many_tag_cr += xml_tag_cr.format(address=address, str_option=str_option, val=val)
        if tag_cr:
            many_tag_row += xml_tag_row.format(row_num=rowID,
                                               row_span='1:'+str(ws_size[0]),
                                               many_tag_cr=many_tag_cr)

    # not 100% what uid does
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
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst uniqueCount="{sharedString_len}" count="{sharedString_len}" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{many_tag_si}
</sst>'''.replace('\n', '\r\n')

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
    # inserts: many_tag_sheets, tag_sharedStrings, tag_calcChain
    xml_base = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{many_tag_sheets}
{tag_sharedStrings}
{tag_calcChain}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''.replace('\n', '\r\n')


    xml_tag_sheet = '<Override PartName="/xl/worksheets/sheet{shID}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\r\n'

    xml_tag_sharedStrings = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\r\n'

    xml_tag_calcChain = '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>\r\n'

    many_tag_sheets = ''
    for shID, _ in enumerate(db.ws_names, 1):
        many_tag_sheets += xml_tag_sheet.format(shID=shID)

    if db._sharedStrings:
        tag_sharedStrings = xml_tag_sharedStrings
    else:
        tag_sharedStrings = ''

    # TODO: once formulas as supported, change this
    tag_calcChain = ''

    rv = xml_base.format(many_tag_sheets=many_tag_sheets,
                         tag_sharedStrings=tag_sharedStrings,
                         tag_calcChain=tag_calcChain)

    return rv

