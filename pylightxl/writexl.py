












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
# inserts: num_sheets, tag_sheets
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
            {tag_sheets}
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
xml_base_2_tag_sheet = '<vt:lpstr>{sheet_name}</vt:lpstr>\r\n'

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
# inserts: tag_sheets, tag_sharedStrings, tag_calcChain
# sheets first for rId# then theme > styles > sharedStrings > calcChain
xml_base_4 = \
'''
<?xml version="1.0" encoding="UTF-8" standalone="true"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    {tag_sheets}
    <Relationship Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Id="rId2"/>
    <Relationship Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Id="rId3"/>
    {tag_sharedStrings}
    {tag_calcChain}
</Relationships>
'''

# location: single tag_sheet insert for xml_base_4
# inserts: sheet_num
xml_base_4_1 = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'

# location: sharedStrings insert for xml_base_4
# inserts: ID
xml_base_4_2 = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId{ID}"/>'

# location: calcChain insert for xml_base_4
# inserts: ID
xml_base_4_3 = '<Relationship Target="calcChain.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Id="rId{ID}"/>'

