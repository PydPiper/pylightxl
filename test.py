import zipfile
import re

file_zip = zipfile.ZipFile('Book2.xlsx')

SHEET_NAMES = []

# actual workbook sheet names are stored in the workbook.xml
with file_zip.open('xl/workbook.xml', 'r') as f:
    text = f.read()
    # convert binary to unicode
    text = text.decode()
    tag_sheet = re.compile(r'sheet name=.*r:id')
    sheet_section = tag_sheet.findall(text)[0]
    # this will find something like:
    # ['sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="sh2" sheetId="2" r:id']
    sheet_text_lines = sheet_section.split('/>')
    for sheet_text_line in sheet_text_lines:
        # split sheet line on '"' will result with: ['sheet name=','Sheet1', 'sheetId=', '1', 'r:id=', 'rId1']
        # simply index to 1 to get the sheet name: Sheet1
        SHEET_NAMES.append(sheet_text_line.split('"')[1])

# the .zip file worksheets are all start Sheet1, Sheet2... names, pull out all the files names that contain "sheet"
zip_sheetnames = [name for name in file_zip.NameToInfo.keys() if 'sheet' in name]

DATABASE = {}

for sheetname in SHEET_NAMES:
    DATABASE.update({sheetname: {}})

for i, zip_sheetname in enumerate(zip_sheetnames):
    with file_zip.open(zip_sheetname, 'r') as f:
        text = f.read()
        text = text.decode()
        tag_sheetdata = re.compile(r'\<sheetData.*sheetData\>')
        sheetdata_section = tag_sheetdata.findall(text)[0]
        tag_cr = re.compile(r'<c r=')
        parts = tag_cr.split(sheetdata_section)[1:]
        for part in parts:
            subparts = part.split('>')
            cell = subparts[0]
            # remove: </v
            val = subparts[2][:-3]
            existing_datadict = DATABASE[SHEET_NAMES[i]]
            existing_datadict.update({cell: val})
            DATABASE.update({SHEET_NAMES[i]: existing_datadict})
        pass

# convert binary to string


pass
