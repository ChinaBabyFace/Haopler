import xlrd
import xlwt

workbook = xlwt.Workbook('utf-8')
worksheet = workbook.add_sheet('2019级8班')
worksheet.write_merge(0, 0, 0, 5, "学生线上学习公式表",)
worksheet.write(1, 0, '序号')
worksheet.write(1, 1, '课程及授课教师')
worksheet.write(1, 2, '周次')
worksheet.write(1, 3, '学生姓名')
worksheet.write(1, 4, '教学任务完成情况')
worksheet.write(1, 5, '作业完成情况')

progressExcelPath = 'C:/Users/green/Documents/第三课 第三节 经济全球化与对外开放.xls'
progressData = xlrd.open_workbook(progressExcelPath)
progressSheet = progressData.sheet_by_index(0)


def get_complete_state(name):
    for i in range(progressSheet.nrows):
        cell = progressSheet.cell(i, 2)
        if cell.value == name and str(progressSheet.cell(i, 10).value) == '完成':
            return '已完成'
    return '未完成'


for j in range(resultSheet.nrows):
    if j > 1:
        print(resultSheet.cell(j, 3).value + ":" + get_complete_state(str(resultSheet.cell(j, 3).value)))
