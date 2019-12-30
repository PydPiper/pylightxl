import zipfile
from os.path import isfile


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
        zf.writestr('_rels/.rels', text_rels.encode())

        text_app = new_app_text(db)
        zf.writestr('docProps/app.xml', text_app.encode())

        text_core = new_core_text(db)
        zf.writestr('docProps/core.xml', text_core.encode())

        text_workbookrels = new_workbookrels_text(db)
        zf.writestr('xl/_rels/workbook.xml.rels', text_workbookrels.encode())

        text_workbook = new_workbook_text(db)
        zf.writestr('xl/workbok.xml', text_workbook.encode())

        # TODO: this needs to be in a loop for each sheet
        text_worksheet = new_worksheet_text(db)
        zf.writestr('xl/worksheets/sheet1.xml', text_worksheet.encode())

        text_sharedStrings = new_sharedStrings_text(db)
        zf.writestr('xl/sharedStrings.xml', text_sharedStrings.encode())

    # TODO: convert zip file to xlsx (rename extension)


def new_rels_text(db):

    # location: /_rels/.rels
    # inserts: -
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>
    '''

    return xml_base


def new_app_text(db):
    """
    Returns /docProps/app.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /docProps/app.xml text
    """

    # location: /docProps/app.xml
    # inserts: num_sheets, many_tag_vt
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="true"?>
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
        <Company/>
        <LinksUpToDate>false</LinksUpToDate>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>16.0300</AppVersion>
    </Properties>
    '''

    # location: single tag_sheet insert for xml_base
    # inserts: sheet_name
    tag_vt = '<vt:lpstr>{sheet_name}</vt:lpstr>\r\n'

    # TODO: finish filling in
    rv = xml_base.format(num_sheets='', many_tag_vt='')

    return rv


def new_core_text(db):
    """
    Returns /docProps/core.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: /docProps/core.xml text
    """

    # location: /docProps/core.xml
    # inserts: -
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="true"?>
    <cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
        <dc:creator>pylightxl</dc:creator>
        <cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>
        <dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>
        <dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>
    </cp:coreProperties>
    '''

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
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="true"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        {many_tag_sheets}
        {tag_sharedStrings}
        {tag_calcChain}
    </Relationships>
    '''

    # location: single tag_sheet insert for xml_base
    # inserts: sheet_num
    xml_tag_sheet = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'

    # location: sharedStrings insert for xml_base
    # inserts: ID
    xml_tag_sharedString = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId{ID}"/>\r\n'

    # location: calcChain insert for xml_base
    # inserts: ID
    #   this will be un-used till cell formulas are supported
    # TODO: add support for formulas at a later time (after writer new/existing are working)
    xml_tag_calcChain = '<Relationship Target="calcChain.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Id="rId{ID}"/>\r\n'

    # TODO: finish filling in
    rv = xml_base.format(many_tag_sheets='',
                         tag_sharedStrings='',
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
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
        <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>
        <workbookPr defaultThemeVersion="166925"/>
        <sheets>
            {many_tag_sheets}
        </sheets>
        <calcPr calcId="181029"/>
    </workbook>
    '''

    # location: worksheet tag for xml_base
    # inserts: name, sheet_id, order_id
    #   note id=rId# is the worksheet order, while sheetId is the worksheet true ID, name= is the custom name
    xml_tag_sheet = '<sheet name="{name}" sheetId="{sheet_id}" r:id="rId{order_id}"/>\r\n'

    # TODO: finish filling in
    rv = xml_base.format(many_tag_sheets='')
    return rv


def new_worksheet_text(db):
    """
    Returns xl/worksheets/sheet#.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/worksheets/sheet#.xml text
    """

    # location: xl/worksheets/sheet#.xml
    # inserts: sheet_size_address (ex: A1:B5, if empty then A1), many_tag_row
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{2C7EE24B-C535-494D-AA97-0A61EE84BA40}">
    <dimension ref="{sheet_size_address}"/>
    <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
    <sheetData>
    {many_tag_row}
    </sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    </worksheet>
    '''

    # location: row tag for xml_base_6
    # inserts: row_num, row_span (1:2), many_tag_cr
    xml_tag_row = '<row r="{row_num}" x14ac:dyDescent="0.25" spans="{row_span}">{many_tag_cr}</row>\r\n'

    # location: c r tag for xml_base_6_tag_row
    # inserts: address, str_option (t="s" for sharedStrings or t="str" for formulas), val
    #   currently formulas are unsupported
    # TODO: add support for formulas at a later time (after writer new/existing are working)
    xml_tag_cr = '<c r="{address}" {str_option}><v>{val}</v></c>'

    # TODO: finish filling in
    rv = xml_base.format(sheet_size_address='', many_tag_row='')
    return rv

def new_sharedStrings_text(db):
    """
    Returns xl/sharedStrings.xml text

    :param pylightxl.Database db: database contains sheetnames, and their data
    :return str: xl/sharedStrings.xml text
    """

    # location: xl/sharedStrings.xml
    # inserts: many_tag_si
    xml_base = \
    '''
    <?xml version="1.0" encoding="UTF-8" standalone="true"?>
    <sst uniqueCount="3" count="4" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        {many_tag_si}
    </sst>
    '''

    # location: si tag for xml_base_7
    # inserts: space_preserve (xml:space="preserve"), val
    #   note leading and trailing spaces requires preserve tag: <t xml:space="preserve"> leadingspace</t>
    xml_tag_si = '<si><t {space_preserve}>{val}</t></si>\r\n'

    # TODO: finish filling in
    rv = xml_base.format(many_tag_si='')
    return rv
