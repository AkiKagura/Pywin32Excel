import public_func
import re


def count_words(worksheet):
    # count = 0
    lst = []
    for i in range(2, 2001):  # 第二个参数需要+1
        val = str(worksheet.Cells(i, 'D').Value)
        # a_list = re.findall(r'[\u4e00-\u9fbf\u3040-\u30ff\uff61-\uff9f]', val)
        search = re.match(r'select|insert|update|delete', val)
        if search:
            lst.append(str(i)+' '+val + '\n')
        # count += len(a_list)
        # print(a_list)
    # print(count)
    public_func.write_list('E:\\TestDatas\\out02261743.txt', lst)


def change_cell_color(worksheet):
    for i in range(1, 2000):
        txt = val = str(worksheet.Cells(i, 'D').Value)
        search = re.search(r'.*(type:).*', val)
        if search is not None:
            worksheet.Cells(i, 'D').Interior.ColorIndex = '36'


def hidden_rows(worksheet):
    for i in range(13, 2001):     # 第二个参数需要+1
        val = str(worksheet.Cells(i, 8).Value)
        # worksheet2.Cells(i + 3, 2).Value = i-1
        # worksheet2.Cells(i + 3, 3).Value = 'EnvName'
        # worksheet2.Cells(i + 3, 11).Value = val
        # worksheet2.Cells(i + 3, 12).Value = val
        if not public_func.is_jpn(val):
            # worksheet.Rows(i).Delete()
            worksheet.Rows(i).Hidden = True
            # attach = ',自動,RES,,'
            # lst.append(str(i) + attach + val + '\n')

