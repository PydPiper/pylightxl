












# location: /_rels/.rels
# inserts: -
xml_base_1 = \
'''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'''

# location: /docProps/app.xml
# inserts: num_sheets, many_tag_vt
xml_base_2 = \
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

# location: single tag_sheet insert for xml_base_2
# inserts: sheet_name
xml_base_2_tag_vt = '<vt:lpstr>{sheet_name}</vt:lpstr>\r\n'

# location: /docProps/core.xml
# inserts: -
xml_base_3 = \
'''
<?xml version="1.0" encoding="UTF-8" standalone="true"?>
<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
    <dc:creator>pylightxl</dc:creator>
    <cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>
</cp:coreProperties>
'''

# location: /xl/_rels/workbook.xml.rels
# inserts: many_tag_sheets, tag_sharedStrings, tag_calcChain
# sheets first for rId# then theme > styles > sharedStrings > calcChain
# note that theme and style is not part of the stack. These don't need to be part of the base xml
xml_base_4 = \
'''
<?xml version="1.0" encoding="UTF-8" standalone="true"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    {many_tag_sheets}
    {tag_sharedStrings}
    {tag_calcChain}
</Relationships>
'''

# location: single tag_sheet insert for xml_base_4
# inserts: sheet_num
xml_base_4_1_tag_sheet = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'

# location: sharedStrings insert for xml_base_4
# inserts: ID
xml_base_4_2 = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId{ID}"/>\r\n'

# location: calcChain insert for xml_base_4
# inserts: ID
# this will be un-used till cell formulas are supported
xml_base_4_3 = '<Relationship Target="calcChain.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Id="rId{ID}"/>\r\n'

# location: xl/workbook.xml
# inserts: many_tag_sheets
xml_base_5 = \
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

# location: worksheet tag for xml_base_5
# inserts: name, sheet_id, order_id
# note id=rId# is the worksheet order, while sheetId is the worksheet true ID, name= is the custom name
xml_base_5_tag_sheet = '<sheet name="{name}" sheetId="{sheet_id}" r:id="rId{order_id}"/>\r\n'


# location: xl/worksheets/sheet1.xml
# inserts: sheet_size_address (ex: A1:B5, if empty then A1), sheetData
xml_base_6 = \
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
xml_base_6_tag_row = '<row r="{row_num}" x14ac:dyDescent="0.25" spans="{row_span}">{many_tag_cr}</row>\r\n'

# location: c r tag for xml_base_6_tag_row
# inserts: address, str_option (t="s" for sharedStrings or t="str" for formulas), val
# currently formulas are unsupported
xml_base_6_tag_cr = '<c r="{address}" {str_option}><v>{val}</v></c>'


# location: xl/sharedStrings.xml
# inserts:
xml_base_7 = \
'''
<?xml version="1.0" encoding="UTF-8" standalone="true"?>
<sst uniqueCount="3" count="4" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    {many_tag_si}
</sst>
'''

# location: si tag for xml_base_7
# inserts: space_preserve (xml:space="preserve"), val
# note leading and trailing spaces requires preserve tag: <t xml:space="preserve"> leadingspace</t>
xml_base_7_tag_si = '<si><t {space_preserve}>{val}</t></si>\r\n'


