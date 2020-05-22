import os
import xlrd
import xlwt
from xlutils.copy import copy

weekId = '3'
resultPath = 'C:/shark/result_save.xlsx'
templatePath = 'C:/shark/result_temp.xlsx'
progressPath = 'C:/shark/class_job.xls'
classPath = 'C:/shark/class_room.xlsx'

if os.path.exists(resultPath):
    os.remove(resultPath)

templateData = xlrd.open_workbook(templatePath, True)
progressData = xlrd.open_workbook(progressPath)
classData = xlrd.open_workbook(classPath)
templateSheet = templateData.sheet_by_index(0)
progressSheet = progressData.sheet_by_index(0)
classSheet = classData.sheet_by_index(0)
resultData = copy(templateData)
resultSheet = resultData.get_sheet(0)


def get_borders():
    borders = xlwt.Borders()  # Create Borders
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    return borders


white_pattern = xlwt.Pattern()  # Create the Pattern
white_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
white_pattern.pattern_fore_colour = 1  # May be: 8 through 63. 0 = Black, 1 = White......
white_grid_style = xlwt.XFStyle()  # Create the Pattern
white_grid_style.pattern = white_pattern  # Add Pattern to Style
white_grid_style.borders = get_borders()

yellow_pattern = xlwt.Pattern()  # Create the Pattern
yellow_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
yellow_pattern.pattern_fore_colour = 5  # May be: 8 through 63. 0 = Black, 1 = White......
yellow_grid_style = xlwt.XFStyle()  # Create the Pattern
yellow_grid_style.pattern = yellow_pattern  # Add Pattern to Style
yellow_grid_style.borders = get_borders()


def get_complete_state(state):
    if state == '完成':
        return '已完成'
    return '未完成'


def get_complete_color(state):
    if state == '完成':
        return yellow_grid_style
    return white_grid_style


def get_class_room_complete_state(student_id):
    # print(student_id)
    for k in range(classSheet.nrows):
        if str(classSheet.cell(k, 1).value) == str(student_id):
            percent = str(classSheet.cell(k, 6).value)
            print(">>" + str(float(percent[0:len(percent) - 1]) >= 100))
            return float(percent[0:len(percent) - 1]) >= 100
    return False


get_class_room_complete_state("2019020941")

for i in range(progressSheet.nrows):
    if i > 1:
        resultSheet.write(i, 0, str(i - 1), white_grid_style)
        resultSheet.write(i, 1, '郝东泽', white_grid_style)
        resultSheet.write(i, 2, weekId, white_grid_style)
        resultSheet.write(i, 3, progressSheet.cell(i, 2).value, white_grid_style)

        flag = '未完成'
        if get_class_room_complete_state(progressSheet.cell(i, 1).value):
            flag = '完成'

        resultSheet.write(i, 4, get_complete_state(flag), get_complete_color(flag))
        resultSheet.write(i, 5, get_complete_state(str(progressSheet.cell(i, 10).value)),
                          get_complete_color(str(progressSheet.cell(i, 10).value)))

        resultSheet.col(0).width = 1500
        resultSheet.col(1).width = 4000
        resultSheet.col(2).width = 1500
        resultSheet.col(3).width = 3000
        resultSheet.col(4).width = 5000
        resultSheet.col(5).width = 5000

        resultData.save(resultPath)
