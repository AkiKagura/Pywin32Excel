import win32com.client
import sheet
import book


if __name__ == '__main__':
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = 0  # 后台运行
    excel.DisplayAlerts = 0  # 不显示，不警告

    # change
    path = 'D:\\MES\\02_SQL差異対応仕様書\\OraclePostgresql SQL差異対応仕様書.xlsx'
    path2 = 'E:\\TestDatas\\MainGATE／IMエンジニアリングツール.xlsx'
    sheet_name = 'Environment'
    # optional
    path_from = 'D:\\MES0205\\01_doc\\05_実装段階\\辞書\\backup\\JISYO_ENV\\POEM_env_Bikou.xlsm'
    name = 'Bikou'

    # open book or sheet
    workbook = excel.Workbooks.Open(path2)
    worksheet = workbook.Worksheets(sheet_name)
    # if needs another workbook, add here

    # func
    sheet.change_cell_color(worksheet)

    # save and close
    workbook.Save()   # if need
    workbook.Close()

    excel.Quit()


