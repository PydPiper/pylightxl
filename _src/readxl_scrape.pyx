import re

def scrape(f, dict sharedString):

    cdef int sample_size
    cdef str text_buff
    cdef dict data
    cdef list match
    cdef str first_match, cell_address, cell_val, next_buff

    data = {}

    sample_size = 10000

    re_cr_tag = re.compile(r'(?<=<c r=)(.+?)(?=</c>)')
    re_cell_val = re.compile(r'(?<=<v>)(.*)(?=</v>)')

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

        while True:
            if match:
                first_match = match.pop(0)
                cell_address = first_match.split('"')[1]
                val_common_str = True if 't="s"' in first_match else False

                cell_val = re_cell_val.findall(first_match)[0]

                if val_common_str:
                    cell_val = sharedString[int(cell_val)]

                data.update({cell_address: cell_val})
            else:
                # only carry forward the reminder unmatched text

                text_buff = re_cr_tag.split(text_buff)[-1]

                next_buff = f.read(sample_size).decode()
                text_buff += next_buff

                break

        if not next_buff:
            break


    return data