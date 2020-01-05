# standard lib imports
from unittest import TestCase

# python27 handling
try:
    ModuleNotFoundError
except NameError:
    ModuleNotFoundError = ImportError

# local lib imports
try:
    from pylightxl.writexl import writexl, new_rels_text, new_app_text, new_core_text, \
        new_workbookrels_text, new_workbook_text, new_worksheet_text, new_sharedStrings_text, \
        new_content_types_text
    from pylightxl.database import Database, address2index, index2address
except ModuleNotFoundError:
    import os, sys

    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('test_writexl'), '..')))

    from pylightxl.writexl import writexl, new_rels_text, new_app_text, new_core_text, \
        new_workbookrels_text, new_workbook_text, new_worksheet_text, new_sharedStrings_text, \
        new_content_types_text
    from pylightxl.database import Database, address2index, index2address


class test_write_new(TestCase):

    def test_rels_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\r\n' \
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\r\n' \
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\r\n' \
                   '</Relationships>'
        self.assertEqual(new_rels_text(None), xml_base)

    def test_app_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
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
        tag_vt = '<vt:lpstr>{sheet_name}</vt:lpstr>\r\n'

        many_tag_vt = tag_vt.format(sheet_name='Sheet1') + \
                      tag_vt.format(sheet_name='Sheet2') + \
                      tag_vt.format(sheet_name='Sheet3') + \
                      tag_vt.format(sheet_name='Sheet4') + \
                      tag_vt.format(sheet_name='Sheet5') + \
                      tag_vt.format(sheet_name='Sheet6') + \
                      tag_vt.format(sheet_name='Sheet7') + \
                      tag_vt.format(sheet_name='Sheet8') + \
                      tag_vt.format(sheet_name='Sheet9') + \
                      tag_vt.format(sheet_name='Sheet10')

        db = Database()
        db.add_ws('Sheet1',{})
        db.add_ws('Sheet2',{})
        db.add_ws('Sheet3',{})
        db.add_ws('Sheet4',{})
        db.add_ws('Sheet5',{})
        db.add_ws('Sheet6',{})
        db.add_ws('Sheet7',{})
        db.add_ws('Sheet8',{})
        db.add_ws('Sheet9',{})
        db.add_ws('Sheet10',{})
        self.assertEqual(new_app_text(db), xml_base.format(num_sheets=10, many_tag_vt=many_tag_vt))

    def test_core_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">\r\n' \
                   '<dc:creator>pylightxl</dc:creator>\r\n' \
                   '<cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>\r\n' \
                   '<dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>\r\n' \
                   '<dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>\r\n' \
                   '</cp:coreProperties>'

        self.assertEqual(new_core_text(None), xml_base)

    def test_workbookrels_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                   '{many_tag_sheets}\r\n' \
                   '{tag_sharedStrings}\r\n' \
                   '</Relationships>'
        xml_tag_sheet = '<Relationship Target="worksheets/sheet{sheet_num}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId{sheet_num}"/>\r\n'
        tag_sharedStrings = '<Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId11"/>\r\n'

        many_tag_sheets = xml_tag_sheet.format(sheet_num=1) + \
                          xml_tag_sheet.format(sheet_num=2) + \
                          xml_tag_sheet.format(sheet_num=3) + \
                          xml_tag_sheet.format(sheet_num=4) + \
                          xml_tag_sheet.format(sheet_num=5) + \
                          xml_tag_sheet.format(sheet_num=6) + \
                          xml_tag_sheet.format(sheet_num=7) + \
                          xml_tag_sheet.format(sheet_num=8) + \
                          xml_tag_sheet.format(sheet_num=9) + \
                          xml_tag_sheet.format(sheet_num=10)

        db = Database()
        db.add_ws('Sheet1',{})
        db.add_ws('Sheet2',{})
        db.add_ws('Sheet3',{})
        db.add_ws('Sheet4',{})
        db.add_ws('Sheet5',{})
        db.add_ws('Sheet6',{})
        db.add_ws('Sheet7',{})
        db.add_ws('Sheet8',{})
        db.add_ws('Sheet9',{})
        db.add_ws('Sheet10',{})
        # test without sharedStrings
        self.assertEqual(new_workbookrels_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=''))
        # test with sharedStrings in db
        db._sharedStrings = ['text']
        self.assertEqual(new_workbookrels_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=tag_sharedStrings))

    def test_workbook_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">\r\n' \
                   '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>\r\n' \
                   '<workbookPr defaultThemeVersion="166925"/>\r\n' \
                   '<sheets>\r\n' \
                   '{many_tag_sheets}\r\n' \
                   '</sheets>\r\n' \
                   '<calcPr calcId="181029"/>\r\n' \
                   '</workbook>'

        xml_tag_sheet = '<sheet name="{sheet_name}" sheetId="{ref_id}" r:id="rId{ref_id}"/>\r\n'

        many_tag_sheets = xml_tag_sheet.format(sheet_name='Sheet1',order_id=1,ref_id=1) + \
                          xml_tag_sheet.format(sheet_name='Sheet2',order_id=2,ref_id=2) + \
                          xml_tag_sheet.format(sheet_name='Sheet3',order_id=3,ref_id=3) + \
                          xml_tag_sheet.format(sheet_name='Sheet4',order_id=4,ref_id=4) + \
                          xml_tag_sheet.format(sheet_name='Sheet5',order_id=5,ref_id=5) + \
                          xml_tag_sheet.format(sheet_name='Sheet6',order_id=6,ref_id=6) + \
                          xml_tag_sheet.format(sheet_name='Sheet7',order_id=7,ref_id=7) + \
                          xml_tag_sheet.format(sheet_name='Sheet8',order_id=8,ref_id=8) + \
                          xml_tag_sheet.format(sheet_name='Sheet9',order_id=9,ref_id=9) + \
                          xml_tag_sheet.format(sheet_name='Sheet10',order_id=10,ref_id=10)

        db = Database()
        db.add_ws('Sheet1',{})
        db.add_ws('Sheet2',{})
        db.add_ws('Sheet3',{})
        db.add_ws('Sheet4',{})
        db.add_ws('Sheet5',{})
        db.add_ws('Sheet6',{})
        db.add_ws('Sheet7',{})
        db.add_ws('Sheet8',{})
        db.add_ws('Sheet9',{})
        db.add_ws('Sheet10',{})
        self.assertEqual(new_workbook_text(db), xml_base.format(many_tag_sheets=many_tag_sheets))

    def test_worksheet_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
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

        xml_tag_row = '<row r="{row_num}" x14ac:dyDescent="0.25" spans="1:{num_of_cr_tags}">{many_tag_cr}</row>\r\n'
        xml_tag_cr = '<c r="{address}" {str_option}>{tag_formula}<v>{val}</v></c>'

        # test all input types
        many_tag_cr_row1 = xml_tag_cr.format(address='A1', str_option='', tag_formula='', val=1) + \
                           xml_tag_cr.format(address='B1', str_option='t="s"', tag_formula='', val=0) + \
                           xml_tag_cr.format(address='C1', str_option='t="str"', tag_formula='<f>A1+2</f>', val='"pylightxl - open excel file and save it for formulas to calculate"')
        # test scarce, and repeat text value
        many_tag_cr_row3 = xml_tag_cr.format(address='A3', str_option='t="s"', tag_formula='', val=0) + \
                           xml_tag_cr.format(address='C3', str_option='t="s"', tag_formula='', val=1)

        many_tag_row = xml_tag_row.format(row_num=1,num_of_cr_tags=3,many_tag_cr=many_tag_cr_row1) + \
                       xml_tag_row.format(row_num=3,num_of_cr_tags=2,many_tag_cr=many_tag_cr_row3)

        uid = '2C7EE24B-C535-494D-AA97-0A61EE84BA40'
        sizeAddress = 'A1:C3'

        db = Database()
        db.add_ws('Sheet1', {'A1':1, 'B1':'text1', 'C1':'=A1+2', 'A3':'text1', 'C3':'text2'})

        self.assertEqual(new_worksheet_text(db, 'Sheet1'), xml_base.format(sizeAddress=sizeAddress,
                                                                           uid=uid,
                                                                           many_tag_row=many_tag_row))

    def test_sharedStrings_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<sst uniqueCount="{sharedString_len}" count="{sharedString_len}" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\r\n' \
                   '{many_tag_si}\r\n' \
                   '</sst>'
        xml_tag_si = '<si><t {space_preserve}>{val}</t></si>\r\n'

        many_tag_si = xml_tag_si.format(space_preserve='',val='text1') + \
                      xml_tag_si.format(space_preserve='',val='text2') + \
                      xml_tag_si.format(space_preserve='xml:space="preserve"',val=' text3') + \
                      xml_tag_si.format(space_preserve='',val='text4') + \
                      xml_tag_si.format(space_preserve='xml:space="preserve"',val='text5 ') + \
                      xml_tag_si.format(space_preserve='xml:space="preserve"',val=' text6 ')

        db = Database()
        db.add_ws('Sheet1', {'A1': 'text1',
                             'A2': 'text2',
                             'A3': ' text3',
                             'A6': 'text4',
                             })

        db.add_ws('Sheet2', {'A4': 'text5 ',
                             'A5': ' text6 ',
                             'A6': 'text4',
                             })

        # process the sharedStrings, see dev note why this is done this way inside new_worksheet_text
        _ = new_worksheet_text(db, 'Sheet1')
        _ = new_worksheet_text(db, 'Sheet2')

        self.assertEqual(new_sharedStrings_text(db), xml_base.format(sharedString_len=6 ,many_tag_si=many_tag_si))

    def test_new_content_types_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
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

        many_tag_sheets = xml_tag_sheet.format(sheet_id=1) + \
                          xml_tag_sheet.format(sheet_id=2) + \
                          xml_tag_sheet.format(sheet_id=3) + \
                          xml_tag_sheet.format(sheet_id=4) + \
                          xml_tag_sheet.format(sheet_id=5) + \
                          xml_tag_sheet.format(sheet_id=6) + \
                          xml_tag_sheet.format(sheet_id=7) + \
                          xml_tag_sheet.format(sheet_id=8) + \
                          xml_tag_sheet.format(sheet_id=9) + \
                          xml_tag_sheet.format(sheet_id=10)

        db = Database()
        db.add_ws('Sheet1', {})
        db.add_ws('Sheet2', {})
        db.add_ws('Sheet3', {})
        db.add_ws('Sheet4', {})
        db.add_ws('Sheet5', {})
        db.add_ws('Sheet6', {})
        db.add_ws('Sheet7', {})
        db.add_ws('Sheet8', {})
        db.add_ws('Sheet9', {})
        db.add_ws('Sheet10', {})

        # test without and sharedStrings in db
        self.assertEqual(new_content_types_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=''))

        # test with and sharedStrings in db
        db._sharedStrings = ['text']
        self.assertEqual(new_content_types_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=xml_tag_sharedStrings))



