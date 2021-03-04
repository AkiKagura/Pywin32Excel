import public_func


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
    public_func.write_excel(out_path, out_list, out_list2)


def copy_sheet(workbook, workbook2, name):
    worksheet = workbook2.Worksheets("Translate")
    worksheet2 = workbook.Worksheets('Sheet1')
    worksheet2.Copy(worksheet)
    workbook.Worksheets('Translate').Name = name


# excel functions
def write_excel(workbook, lst1, lst2):
    ws = workbook.Worksheets("Sheet1")
    lst_len = len(lst1)
    for i in range(0, lst_len):
        ws.Cells(i + 1, 'A').Value = lst1[i]
        ws.Cells(i + 1, 'E').Value = lst2[i]




