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

SHARED_STRINGS = {}

# check for sharedString.xml file that contains string values of cells
if 'xl/sharedStrings.xml' in file_zip.NameToInfo.keys():
    with file_zip.open('xl/sharedStrings.xml') as f:
        text = f.read()
        text = text.decode()

        tag_t = re.compile(r'<t>(.*?)</t>')
        tag_t_vals = tag_t.findall(text)
        for i, val in enumerate(tag_t_vals):
            SHARED_STRINGS.update({i: val})

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
        tag_cr_lines = tag_cr.split(sheetdata_section)[1:]
        for tag_cr_line in tag_cr_lines:
            # pull out cell address and test if it's a string cell that needs lookup
            re_cell_address = re.compile(r'[^<r c="][^"]+')
            finding_cell_address = re_cell_address.findall(tag_cr_line)
            cell_address = finding_cell_address[0]
            cell_string = True if 't="s"' in tag_cr_line else False

            re_cell_val = re.compile(r'(?<=<v>)(.*)(?=</v>)')
            cell_val = re_cell_val.findall(tag_cr_line)[0]

            if cell_string is True:
                cell_val = SHARED_STRINGS[int(cell_val)]

            # add to db
            existing_datadict = DATABASE[SHEET_NAMES[i]]
            existing_datadict.update({cell_address: cell_val})
            DATABASE.update({SHEET_NAMES[i]: existing_datadict})
        pass

# convert binary to string

print(DATABASE)
pass
