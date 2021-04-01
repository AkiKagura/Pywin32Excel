import public_func
import sheet


def count_sheets(workbook):
    # out_path = 'E:\\TestDatas\\sql0225.xlsx'
    sheet_num = workbook.Sheets.Count
    out_list = []
    # out_list2 = []
    menu_ws = workbook.Worksheets(3)
    for i in range(4, sheet_num + 1 - 2):
        ws = workbook.Worksheets(i)
        out_list.append(ws.Name)
        menu_ws.Cells(i+6, 'D').Value = ws.Name
        '''for j in range(12, 51):
            val = str(ws.Cells(j, 'D').Value)
            val2 = str(ws.Cells(j, 'AG').Value)
            if ws.Cells(j, 'D').Font.Bold == 1:
                out_list.append(val)
                out_list2.append(val2)'''
    # public_func.write_excel(out_path, out_list, out_list2)


# copy sheet1 before sheet2 and rename it
def copy_sheet(workbook, to_copy, workbook2, copy_before, name):
    to_copy_sheet = workbook.Worksheets(to_copy)
    copy_before_sheet = workbook2.Worksheets(copy_before)
    to_copy_sheet.Copy(copy_before_sheet)
    new_sheet = workbook.Worksheets(workbook.Sheets.Count - 3)
    new_sheet.Name = name


# excel functions
def write_excel(workbook, lst1, lst2):
    ws = workbook.Worksheets("Sheet1")
    lst_len = len(lst1)
    for i in range(0, lst_len):
        ws.Cells(i + 1, 'A').Value = lst1[i]
        ws.Cells(i + 1, 'E').Value = lst2[i]


def show_all_sheets_rows(workbook):
    sheet_num = workbook.Sheets.Count
    for i in range(1, sheet_num + 1):
        ws = workbook.Worksheets(i)
        sheet.show_rows(ws)












