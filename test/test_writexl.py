# standard lib imports
from unittest import TestCase
import os, sys, shutil, io

if sys.version_info[0] >= 3:
    unicode = str
    FileNotFoundError = IOError
    PermissionError = Exception
else:
    ModuleNotFoundError = ImportError

try:
    from pylightxl import pylightxl as xl
except ModuleNotFoundError:
    sys.path.append('..')
    from pylightxl import pylightxl as xl


if 'test' in os.listdir('.'):
    # running from top level
    os.chdir('./test')


class TestWritexlNew(TestCase):

    def test_rels_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n' \
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\r\n' \
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\r\n' \
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\r\n' \
                   '</Relationships>'
        self.assertEqual(xl.writexl_new_rels_text(None), xml_base)

    def test_app_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">\r\n' \
                   '<Application>Microsoft Excel</Application>\r\n' \
                   '<DocSecurity>0</DocSecurity>\r\n' \
                   '<ScaleCrop>false</ScaleCrop>\r\n' \
                   '<HeadingPairs>\r\n' \
                   '<vt:vector baseType="variant" size="{vector_size}">\r\n' \
                   '<vt:variant>\r\n' \
                   '<vt:lpstr>Worksheets</vt:lpstr>\r\n' \
                   '</vt:variant>\r\n' \
                   '<vt:variant>\r\n' \
                   '<vt:i4>{ws_size}</vt:i4>\r\n' \
                   '</vt:variant>\r\n' \
                   '{variant_tag_nr}' \
                   '</vt:vector>\r\n' \
                   '</HeadingPairs>\r\n' \
                   '<TitlesOfParts>\r\n' \
                   '<vt:vector baseType="lpstr" size="{vt_size}">\r\n' \
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

        db = xl.Database()
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
        self.assertEqual(xl.writexl_new_app_text(db), xml_base.format(vector_size=2,
                                                                      ws_size=10,
                                                                      variant_tag_nr='',
                                                                      vt_size=10,
                                                                      many_tag_vt=many_tag_vt))

        db.add_nr('range1', 'Sheet1', 'A1')
        db.add_nr('range2', 'Sheet1', 'A2:A5')

        variant_tag_nr = '<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant>\r\n' \
                         '<vt:variant><vt:i4>2</vt:i4></vt:variant>\r\n'

        # python 2 does not keep dict order
        if sys.version_info[0] >= 3:
            many_tag_vt += '<vt:lpstr>range1</vt:lpstr>\r\n'
            many_tag_vt += '<vt:lpstr>range2</vt:lpstr>\r\n'
        else:
            many_tag_vt += '<vt:lpstr>range2</vt:lpstr>\r\n'
            many_tag_vt += '<vt:lpstr>range1</vt:lpstr>\r\n'

        self.assertEqual(xl.writexl_new_app_text(db), xml_base.format(vector_size=4,
                                                                      ws_size=10,
                                                                      variant_tag_nr=variant_tag_nr,
                                                                      vt_size=12,
                                                                      many_tag_vt=many_tag_vt))

    def test_core_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">\r\n' \
                   '<dc:creator>pylightxl</dc:creator>\r\n' \
                   '<cp:lastModifiedBy>pylightxl</cp:lastModifiedBy>\r\n' \
                   '<dcterms:created xsi:type="dcterms:W3CDTF">2019-12-27T01:35:28Z</dcterms:created>\r\n' \
                   '<dcterms:modified xsi:type="dcterms:W3CDTF">2019-12-27T01:35:39Z</dcterms:modified>\r\n' \
                   '</cp:coreProperties>'

        self.assertEqual(xl.writexl_new_core_text(None), xml_base)

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

        db = xl.Database()
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
        self.assertEqual(xl.writexl_new_workbookrels_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=''))
        # test with sharedStrings in db
        db._sharedStrings = ['text']
        self.assertEqual(xl.writexl_new_workbookrels_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=tag_sharedStrings))

    def test_workbook_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">\r\n' \
                   '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="22228"/>\r\n' \
                   '<workbookPr defaultThemeVersion="166925"/>\r\n' \
                   '<sheets>\r\n' \
                   '{many_tag_sheets}\r\n' \
                   '</sheets>\r\n' \
                   '{xml_namedrange}' \
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

        xml_namedrange = '<definedNames><definedName name="{}">{}</definedName>\r\n'.format('range1', 'Sheet1!A1') + \
                      '<definedName name="{}">{}</definedName>\r\n</definedNames>\r\n'.format('range2', 'Sheet2!A1:C3')

        db = xl.Database()
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
        self.assertEqual(xl.writexl_new_workbook_text(db),
                         xml_base.format(many_tag_sheets=many_tag_sheets, xml_namedrange=''))

        db.add_nr('range1', 'Sheet1', 'A1')
        db.add_nr('range2', 'Sheet2', 'A1:C3')
        try:
            self.assertEqual(xl.writexl_new_workbook_text(db),
                             xml_base.format(many_tag_sheets=many_tag_sheets, xml_namedrange=xml_namedrange))
        except:
            # python 2 does not keep dict order
            xml_namedrange = '<definedNames><definedName name="{}">{}</definedName>\r\n'.format('range2',
                                                                                                'Sheet2!A1:C3') + \
                             '<definedName name="{}">{}</definedName>\r\n</definedNames>\r\n'.format(
                                 'range1', 'Sheet1!A1')
            self.assertEqual(xl.writexl_new_workbook_text(db),
                             xml_base.format(many_tag_sheets=many_tag_sheets, xml_namedrange=xml_namedrange))

    def test_worksheet_text(self):
        xml_base = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
                   '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{uid}">\r\n' \
                   '<dimension ref="{sizeAddress}"/>\r\n' \
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
                           '<c r="{address}">{tag_formula}</c>'.format(address='C1', tag_formula='<f>A1+2</f>') + \
                           xml_tag_cr.format(address='D1', str_option='t="s"', tag_formula='', val=1)
            # test scarce, and repeat text value
        many_tag_cr_row3 = xml_tag_cr.format(address='A3', str_option='t="s"', tag_formula='', val=0) + \
                           xml_tag_cr.format(address='C3', str_option='t="s"', tag_formula='', val=2)

        many_tag_row = xml_tag_row.format(row_num=1,num_of_cr_tags=4,many_tag_cr=many_tag_cr_row1) + \
                       xml_tag_row.format(row_num=3,num_of_cr_tags=2,many_tag_cr=many_tag_cr_row3)

        uid = '2C7EE24B-C535-494D-AA97-0A61EE84BA40'
        sizeAddress = 'A1:D3'

        db = xl.Database()
        db.add_ws('Sheet1', {})
        db.ws('Sheet1').update_address('A1', 1)
        db.ws('Sheet1').update_address('B1', 'text1')
        db.ws('Sheet1').update_address('C1', '=A1+2')
        db.ws('Sheet1').update_address('D1', 'B1&"_"&"two"')
        db.ws('Sheet1').update_address('A3', 'text1')
        db.ws('Sheet1').update_address('C3', 'text2')

        self.assertEqual(xl.writexl_new_worksheet_text(db, 'Sheet1'), xml_base.format(sizeAddress=sizeAddress,
                                                                           uid=uid,
                                                                           many_tag_row=many_tag_row))
        #TODO: add checks for sharedStrings

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

        db = xl.Database()
        db.add_ws('Sheet1', {'A1': {'v': 'text1', 'f': '', 's': ''},
                             'A2': {'v': 'text2', 'f': '', 's': ''},
                             'A3': {'v': ' text3', 'f': '', 's': ''},
                             'A6': {'v': 'text4', 'f': '', 's': ''},
                             })

        db.add_ws('Sheet2', {'A4': {'v': 'text5 ', 'f': '', 's': ''},
                             'A5': {'v': ' text6 ', 'f': '', 's': ''},
                             'A6': {'v': 'text4', 'f': '', 's': ''},
                             })

        # process the sharedStrings, see dev note why this is done this way inside new_worksheet_text
        _ = xl.writexl_new_worksheet_text(db, 'Sheet1')
        _ = xl.writexl_new_worksheet_text(db, 'Sheet2')

        self.assertEqual(xl.writexl_new_sharedStrings_text(db), xml_base.format(sharedString_len=6 ,many_tag_si=many_tag_si))

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

        db = xl.Database()
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
        self.assertEqual(xl.writexl_new_content_types_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=''))

        # test with and sharedStrings in db
        db._sharedStrings = ['text']
        self.assertEqual(xl.writexl_new_content_types_text(db), xml_base.format(many_tag_sheets=many_tag_sheets, tag_sharedStrings=xml_tag_sharedStrings))

    def test_openpyxl(self):
        # test that pylightxl is able to write to a openpyxl output excel file (docProps/app.xml) is different than expected
        db = xl.readxl('openpyxl.xlsx')

        xl.writexl(db, 'newopenpyxl.xlsx')


class TestWritexlExisting(TestCase):

    def test_writexl_alt_app_text(self):
        # tests that sheet names and sheet count were updated and named ranges were preserved
        db = xl.Database()
        # existing info in input_app.xml
        db.add_ws('Sheet1')
        db.add_ws('sh2')
        # new info
        db.add_ws('one')
        db.add_ws('two')
        db.add_ws('three')

        text = xl.writexl_alt_app_text(db=db, filepath='input_app.xml')

        if sys.version_info[0] < 3:
            with open('correct_app27.xml', 'r') as f:
                correct_text = f.read()
        elif sys.version_info[1] == 7:
            with open('correct_app37.xml', 'r') as f:
                correct_text = f.read()
        else:
            with open('correct_app3.xml', 'r') as f:
                correct_text = f.read()

        self.assertEqual(correct_text, text)

    def test_writexl_alt_app_text_withNR(self):
        # tests that sheet names and sheet count were updated and named ranges were preserved
        db = xl.Database()
        # existing info in input_app.xml
        db.add_ws('Sheet1')
        db.add_ws('sh2')
        db.add_nr('mylist', 'Sheet1', 'A1')
        db.add_nr('mylist2', 'Sheet1', 'A2')
        # new info
        db.add_ws('one')
        db.add_ws('two')
        db.add_ws('three')

        text = xl.writexl_alt_app_text(db=db, filepath='input_app_withNR.xml')

        if sys.version_info[0] < 3:
            with open('correct_app27_withNR.xml', 'r') as f:
                correct_text = f.read()
        elif sys.version_info[1] == 7:
            with open('correct_app37_withNR.xml', 'r') as f:
                correct_text = f.read()
        else:
            with open('correct_app3_withNR.xml', 'r') as f:
                correct_text = f.read()

        self.assertEqual(correct_text, text)

    def test_writexl_alt_getsheetref(self):
        sheetref = xl.writexl_alt_getsheetref(path_wbrels='input_workbook.xml.rels',
                                              path_wb='input_workbook.xml')

        correct_sheetref = {'rId2': {'sheetId': 2, 'name': 'sh2', 'filename': 'sheet2.xml'},
                            'rId1': {'sheetId': 1, 'name': 'Sheet1', 'filename': 'sheet1.xml'},
                            }

        self.assertEqual(correct_sheetref, sheetref)

    def test_integration_alt_writer(self):
        db = xl.Database()

        # cleanup failed test workbook
        if 'temp_wb.xlsx' in os.listdir('.'):
            os.remove('temp_wb.xlsx')
        if '_pylightxl_temp_wb.xlsx' in os.listdir('.'):
            shutil.rmtree('_pylightxl_temp_wb.xlsx')

        # create the "existing workbook"
        db.add_ws(ws='sh1', data={'A1': {'v':'one', 'f': '', 's': ''},
                                         'A2': {'v':1, 'f': '', 's': ''},
                                         'A3': {'v':1.0, 'f': '', 's': ''},
                                         'A4': {'v':'one', 'f': 'A1', 's': ''},
                                         'A5': {'v':6, 'f': 'A2+5', 's': ''},
                                         'B1': {'v': 'one', 'f': '', 's': ''},
                                         'B2': {'v': 1, 'f': '', 's': ''},
                                         'B3': {'v': 1.0, 'f': '', 's': ''},
                                         'B4': {'v': 'one', 'f': 'A1', 's': ''},
                                         'B5': {'v': 6, 'f': 'A2+5', 's': ''},
                                         })
        db.add_ws(ws='sh2')
        xl.writexl(db, 'temp_wb.xlsx')

        # all changes will be registered as altered xl writer since the filename exists
        db.ws(ws='sh1').update_address('B1', 'two')
        db.ws(ws='sh1').update_address('B2', 2)
        db.ws(ws='sh1').update_address('B3', 2.0)
        # was a formula now a string that looks like a formula
        db.ws(ws='sh1').update_address('B4', 'A1&"_"&"two"')
        db.ws(ws='sh1').update_address('B5', '=A2+10')
        db.ws(ws='sh1').update_address('C6', 'new')

        db.add_ws(ws='sh3')
        db.ws(ws='sh3').update_address('A1', 'one')

        xl.writexl(db, 'temp_wb.xlsx')

        # check the results made it in correctly
        db_alt = xl.readxl(fn='temp_wb.xlsx')

        self.assertEqual([6, 3], db_alt.ws('sh1').size)
        self.assertEqual('one', db_alt.ws('sh1').address('A1'))
        self.assertEqual(1, db_alt.ws('sh1').address('A2'))
        self.assertEqual(1.0, db_alt.ws('sh1').address('A3'))
        self.assertEqual('', db_alt.ws('sh1').address('A4'))
        self.assertEqual('A1', db_alt.ws('sh1')._data['A4']['f'])
        self.assertEqual('', db_alt.ws('sh1').address('A5'))
        self.assertEqual('A2+5', db_alt.ws('sh1')._data['A5']['f'])
        self.assertEqual('two', db_alt.ws('sh1').address('B1'))
        self.assertEqual(2, db_alt.ws('sh1').address('B2'))
        self.assertEqual(2.0, db_alt.ws('sh1').address('B3'))
        self.assertEqual('A1&amp;"_"&amp;"two"', db_alt.ws('sh1').address('B4'))
        self.assertEqual('', db_alt.ws('sh1')._data['B4']['f'])
        self.assertEqual('', db_alt.ws('sh1').address('B5'))
        self.assertEqual('A2+10', db_alt.ws('sh1')._data['B5']['f'])
        self.assertEqual('new', db_alt.ws('sh1').address('C6'))

        self.assertEqual([0, 0], db_alt.ws('sh2').size)
        self.assertEqual('', db_alt.ws('sh2').address('A1'))

        self.assertEqual([1, 1], db_alt.ws('sh3').size)
        self.assertEqual('one', db_alt.ws('sh3').address('A1'))

        # cleanup failed test workbook
        if 'temp_wb.xlsx' in os.listdir('.'):
            os.remove('temp_wb.xlsx')


class TestWriteCSV(TestCase):

    def test_writecsv(self):

        db = xl.Database()
        db.add_ws('sh1')
        db.add_ws('sh2')
        db.ws('sh1').update_index(1,1, 10)
        db.ws('sh1').update_index(1,2, 10.0)
        db.ws('sh1').update_index(1,3, '10.0\n')
        db.ws('sh1').update_index(1,4, True)
        db.ws('sh1').update_index(2,1, 20)
        db.ws('sh1').update_index(2,2, 20.0)
        db.ws('sh1').update_index(2,3, '20.0')
        db.ws('sh1').update_index(2,4, False)
        db.ws('sh1').update_index(3,5, ' ')
        db.ws('sh2').update_index(1,1, 'sh2')


        if 'outcsv_sh1.csv' in os.listdir('.'):
            os.remove('outcsv_sh1.csv')
        if 'outcsv_sh2.csv' in os.listdir('.'):
            os.remove('outcsv_sh2.csv')

        xl.writecsv(db=db, fn='outcsv', delimiter='\t', ws='sh1')

        with open('outcsv_sh1.csv', 'r') as f:
            lines = []
            while True:
                line = f.readline()

                if not line:
                    break

                line = line.replace('\n', '').replace('\r', '')

                lines.append(line.split('\t'))

        self.assertEqual(['10', '10.0', '10.0', 'True', ''], lines[0])
        self.assertEqual(['20', '20.0', '20.0', 'False', ''], lines[1])
        self.assertEqual(['', '', '', '', ' '], lines[2])

        if 'outcsv_sh1.csv' in os.listdir('.'):
            os.remove('outcsv_sh1.csv')
        if 'outcsv_sh2.csv' in os.listdir('.'):
            os.remove('outcsv_sh2.csv')

        xl.writecsv(db=db, fn='outcsv')

        self.assertTrue('outcsv_sh1.csv' in os.listdir('.'))
        self.assertTrue('outcsv_sh2.csv' in os.listdir('.'))

        if 'outcsv_sh1.csv' in os.listdir('.'):
            os.remove('outcsv_sh1.csv')
        if 'outcsv_sh2.csv' in os.listdir('.'):
            os.remove('outcsv_sh2.csv')

        f = io.StringIO()

        xl.writecsv(db=db, fn=f)

        f.seek(0)
        self.assertEqual('10,10.0,10.0,True,\n', f.readline())
        self.assertEqual('20,20.0,20.0,False,\n', f.readline())
        self.assertEqual(',,,, \n', f.readline())
        self.assertEqual('sh2\n', f.readline())
