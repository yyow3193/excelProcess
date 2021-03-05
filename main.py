import os

import xlrd
import xlwt
import copy


def set_style(name, height, bold=True):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name
    font.bold = bold
    #font.color_index = 4
    font.height = height

    style.font = font
    return style


file_name='8.1-8.31日福龙结算单(12)_8.xlsx'
excel_file=os.getcwd()+'\\'+file_name
rdata=xlrd.open_workbook(excel_file)
print('sheets nums:', rdata.nsheets)  # excel sheets 个数

monthindex = file_name.find("_", len(file_name) - 7, len(file_name)) + 1
month = file_name[monthindex]

# 每个人按月汇总

book_everymonth_everypersion = xlwt.Workbook(encoding='utf-8')

name2newsheet = {}
name2totalrow = {}
titlerow = None

# 汇总每个月的每个人
allsheet = book_everymonth_everypersion.add_sheet("all", cell_overwrite_ok=True)


for sheet in rdata.sheets():
    print("open sheet index:", sheet.name)
    for rowindex in range(sheet.nrows):
        if rowindex <= 2:
            if titlerow == None and rowindex == 2:
                titlerow = sheet.row(rowindex)
            continue
        row = sheet.row(rowindex)
        if row[5].value == "":
            continue
        newsheet = None
        if row[5].value in name2newsheet:
            newsheet = name2newsheet[row[5].value]
            name2totalrow[row[5].value][2].value = name2totalrow[row[5].value][2].value + row[2].value
            name2totalrow[row[5].value][3].value = name2totalrow[row[5].value][3].value + row[3].value
            name2totalrow[row[5].value][4].value = name2totalrow[row[5].value][4].value + row[4].value

        else:
            newsheet = book_everymonth_everypersion.add_sheet(row[5].value, cell_overwrite_ok=True)
            name2newsheet[row[5].value] = newsheet

            titlerow = sheet.row(2)
            for colindex in range(len(titlerow)):
                newsheet.write(0, colindex, titlerow[colindex].value)

            name2totalrow[row[5].value] = copy.deepcopy(row)

        if(len(allsheet.get_rows()) <= 0):
            for colindex in range(len(titlerow)):
                allsheet.write(0, colindex, titlerow[colindex].value)

        # style = xlwt.XFStyle()  # 初始化样式
        # font = xlwt.Font()  # 为样式创建字体
        # font.name = 'Times New Roman'
        # font.bold = True  # 黑体
        # font.underline = True  # 下划线
        # font.italic = True  # 斜体字
        # style.font = font  # 设定样式
        #
        # borders = xlwt.Borders()
        # borders.left = xlwt.Borders.THIN
        # borders.left = xlwt.Borders.THIN
        # # NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
        # # THIN： 官方代码中THIN所表示的值为1，边框为实线
        # # 左边框 细线
        # borders.left = 1
        # # 右边框 中细线
        # borders.right = 2
        # # 上边框 虚线
        # borders.top = 3
        # # 下边框 点线
        # borders.bottom = 4
        # # 内边框 粗线
        # borders.diag = 5
        #
        # # 左边框颜色 蓝色
        # borders.left_colour = 0x0C
        # # 右边框颜色 金色
        # borders.right_colour = 0x33
        # # 上边框颜色 绿色
        # borders.top_colour = 0x11
        # # 下边框颜色 红色
        # borders.bottom_colour = 0x0A
        # # 内边框 黄色
        # borders.diag_colour = 0x0D
        # # 定义格式
        # style.borders = borders



        nowrowcount = len(newsheet.get_rows())
        for colindex in range(len(row)):
            newsheet.write(nowrowcount, colindex, row[colindex].value)
            #print(row[colindex].value)

for (k, v) in name2totalrow.items():
    nowrowcount = len(allsheet.get_rows())
    print(v)
    for colindex in range(len(v)):
        allsheet.write(nowrowcount, colindex, v[colindex].value)


everymonth_everypersion_book_name = "book_everymonth_everypersion_" + month + ".xls"
book_everymonth_everypersion.save(everymonth_everypersion_book_name)


# everymonth_everypersion_book_name=os.getcwd()+'\\'+everymonth_everypersion_book_name
# rdata1=xlrd.open_workbook(everymonth_everypersion_book_name)
#
# allsheet = rdata1.sheet_by_name("all")
# print(allsheet.name)
#
# for sheet in rdata.sheets():
#     allrow = None
#     for rowindex in range(sheet.nrows):
#         if(rowindex==0):
#             continue
#
#         for rowindex in range(sheet.nrows):
#             row = sheet.row(rowindex)
#             if (allrow == None):
#                 allrow = copy.deepcopy(row)
#             else:
#                 allrow[2].value = allrow[2].value + row[2].value
#                 allrow[3].value = allrow[3].value + row[3].value
#                 allrow[4].value = allrow[4].value + row[4].value
#                 allrow[9].value = allrow[9].value + row[9].value







