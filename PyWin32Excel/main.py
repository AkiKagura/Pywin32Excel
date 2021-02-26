import win32com.client
import re


# functions
def is_jpn(string):
    for ch in string:
        if u'\u3040' <= ch <= u'\u30ff':
            return True

    return False


def write_list(txt_path, lst):
    f = open(txt_path, 'w', encoding='UTF-8')
    for var in lst:
        f.write(var)
    f.close()


# excel functions
def write_excel(path_out, lst1, lst2):
    wb = excel.Workbooks.Open(path_out)
    ws = wb.Worksheets("Sheet1")
    lst_len = len(lst1)
    for i in range(0, lst_len):
        ws.Cells(i + 1, 'A').Value = lst1[i]
        ws.Cells(i + 1, 'E').Value = lst2[i]
    wb.Save()
    wb.Close()


def hidden_rows(worksheet):
    for i in range(13, 2001):     # 第二个参数需要+1
        val = str(worksheet.Cells(i, 8).Value)
        # worksheet2.Cells(i + 3, 2).Value = i-1
        # worksheet2.Cells(i + 3, 3).Value = 'EnvName'
        # worksheet2.Cells(i + 3, 11).Value = val
        # worksheet2.Cells(i + 3, 12).Value = val
        if not is_jpn(val):
            # worksheet.Rows(i).Delete()
            worksheet.Rows(i).Hidden = True
            # attach = ',自動,RES,,'
            # lst.append(str(i) + attach + val + '\n')


def copy_sheet(path_from, workbook, name):
    workbook2 = excel.Workbooks.Open(path_from)
    worksheet2 = workbook2.Worksheets("Translate")
    worksheet = workbook.Worksheets('Sheet1')
    worksheet2.Copy(worksheet)
    workbook.Worksheets('Translate').Name = name
    workbook2.Close()


def count_words(worksheet):
    count = 0
    for i in range(13, 2001):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 8).Value)
        a_list = re.findall(r'[\u4e00-\u9fbf\u3040-\u30ff\uff61-\uff9f]', val)
        count += len(a_list)
        # print(a_list)
    print(count)


def read_cells(workbook):
    out_path = 'E:\\TestDatas\\sql0225.xlsx'
    sheet_num = workbook.Sheets.Count
    out_list = []
    out_list2 = []
    for i in range(1, sheet_num + 1):
        ws = workbook.Worksheets(i)
        for j in range(12, 51):
            val = str(ws.Cells(j, 'D').Value)
            val2 = str(ws.Cells(j, 'AG').Value)
            if ws.Cells(j, 'D').Font.Bold == 1:
                out_list.append(val)
                out_list2.append(val2)

    print(out_list)
    write_excel(out_path, out_list, out_list2)


if __name__ == '__main__':
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = 0  # 后台运行
    excel.DisplayAlerts = 0  # 不显示，不警告

    # change
    path = 'D:\\MES\\02_SQL差異対応仕様書\\OraclePostgresql SQL差異対応仕様書.xlsx'
    sheet_name = 'Bikou'
    # optional
    path_from = 'D:\\MES0205\\01_doc\\05_実装段階\\辞書\\backup\\JISYO_ENV\\POEM_env_Bikou.xlsm'
    name = 'Bikou'

    # open book or sheet
    workbook = excel.Workbooks.Open(path)
    # worksheet = workbook.Worksheets(sheet_name)

    # func
    # copy_sheet(path_from, workbook, name)
    # count_words(worksheet)
    read_cells(workbook)

    # workbook.Save()   # if need
    workbook.Close()

    excel.Quit()


