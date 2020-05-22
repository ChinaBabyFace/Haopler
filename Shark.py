import xlrd
import xlwt

resultExcelPath = 'C:/Users/green/Documents/result.xlsx'
progressExcelPath = 'C:/Users/green/Documents/第三课 第三节 经济全球化与对外开放.xls'

resultExcelData = xlrd.open_workbook(resultExcelPath)
progressData = xlrd.open_workbook(progressExcelPath)

resultSheet = resultExcelData.sheet_by_index(2)
progressSheet = progressData.sheet_by_index(0)

workbook = xlwt.Workbook('utf-8')
workbook.add


def get_complete_state(name):
    for i in range(progressSheet.nrows):
        cell = progressSheet.cell(i, 2)
        if cell.value == name and str(progressSheet.cell(i, 10).value) == '完成':
            return '已完成'
    return '未完成'


for j in range(resultSheet.nrows):
    if j > 1:
        print(resultSheet.cell(j, 3).value + ":" + get_complete_state(str(resultSheet.cell(j, 3).value)))
