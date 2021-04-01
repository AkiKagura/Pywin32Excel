import public_func
import re


def count_words(worksheet, out_path):
    # count = 0
    lst = []
    for i in range(2, 2001):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 'D').Value)
        # a_list = re.findall(r'[\u4e00-\u9fbf\u3040-\u30ff\uff61-\uff9f]', val)
        # search = re.match(r'select|insert|update|delete', val)
        # if search:
        #     lst.append(str(i)+' '+val + '\n')
        if val is not None:
            lst.append(str(i) + ' ' + val + '\n')
        # count += len(a_list)
        # print(a_list)
    # print(count)
    public_func.write_list(out_path, lst)


# output some cells of a sheet
def sheet_output_words(worksheet, out_path):
    lst = []
    for i in range(2, 2466):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 'D').Value)
        if val != '':
            match = re.search(
                r'()',
                val, re.I)
            if match is not None:
                lst.append('Line ' + str(i) + ': ' + val + '\n')
    public_func.write_list(out_path, lst)


# change the value of some cells
def sheet_change_words(worksheet, out_path):
    for i in range(2, 2466):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 'D').Value)
        match = re.search(r'(insert into )(\S+_lang)( \( )(.*)( \) )(values)( \( )(.*)( \))', val, re.I)
        if match is not None:
            lst = list(match.groups())
            lst[3] += ', language_id'
            lst[7] += ', \'001\''
            worksheet.Cells(i, 'D').Value = ''.join(lst)


# hide rows
def hide_rows(worksheet):
    for i in range(13, 2001):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 8).Value)
        if not public_func.is_jpn(val):
            worksheet.Rows(i).Hidden = True


# show hidden rows
def show_rows(worksheet):
    for i in range(1, 3001):  # 第二个参数需要+1
        if worksheet.Rows(i).Hidden:
            worksheet.Rows(i).Hidden = False


def cell_find(worksheet):
    for i in range(2, 830):
        val = str(worksheet.Cells(i, 'D').Value)
        search = re.search(r'dmi_layout_def|dmi_system_def|dmi_table_create', val)
        if search is not None:
            worksheet.Cells(i, 'D').Interior.ColorIndex = '17'


def sheet_compare(sheet_1, sheet_2):  # 1 is to be changed
    for i in range(2, 2466):
        val1 = str(sheet_1.Cells(i, 'D').Value)
        val2 = str(sheet_2.Cells(i, 'D').Value)
        if val1 != val2:
            sheet_1.Cells(i, 'D').Interior.ColorIndex = '36'


# get text in a range
def get_text(w_sheet, srt_r, srt_c, end_r, end_c):
    str_res = ''
    cell_srt = w_sheet.Cells(srt_r, srt_c)
    cell_end = w_sheet.Cells(end_r, end_c)
    for cell in w_sheet.Range(cell_srt, cell_end):
        if cell.Value is not None:
            str_res += cell.Value
        else:
            str_res += '\n'
    str_res2 = re.sub(r'\n*', ' ', str_res)
    return str_res2


# open sheet and get a dictionary
def get_dictionary(w_sheet):
    sheet_dict = {}
    for j in range(13, 3000):
        word_jpn = str(w_sheet.Cells(j, 'H').Value)
        word_chn = str(w_sheet.Cells(j, 'I').Value)
        # if re.search(r'[\u4e00-\u9fbf\u3040-\u30ff\uff61-\uff9f]', word_jpn) is not None:
        if word_jpn is not None:
            sheet_dict[word_jpn] = word_chn
    return sheet_dict


# open sheet and write into a dictionary
def set_dictionary(w_sheet, sheet_dict):
    for j in range(13, 3000):
        jpn = w_sheet.Cells(j, 'H').Value
        word_jpn = str(jpn)
        if word_jpn in sheet_dict and jpn is not None:
            w_sheet.Cells(j, 'I').Value = sheet_dict[word_jpn]
