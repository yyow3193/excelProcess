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
name2recordsumlist = {}

name2recordlist = {}
titlerow = None

# 汇总每个月的每个人
for sheet in rdata.sheets():
    print("open sheet index:", sheet.name)
    for rowindex in range(sheet.nrows):
        if rowindex <= 2: # 前两行是公司名
            if titlerow == None and rowindex == 2:
                titlerow = sheet.row(rowindex)
            continue
        row = sheet.row(rowindex)
        if row[5].value == "":
            continue

        if row[5].value in name2recordlist:
            recordlist = name2recordlist[row[5].value]
            recordlist.append(row)

            name2recordsumlist[row[5].value][2].value = name2recordsumlist[row[5].value][2].value + row[2].value
            name2recordsumlist[row[5].value][3].value = name2recordsumlist[row[5].value][3].value + row[3].value
            name2recordsumlist[row[5].value][4].value = name2recordsumlist[row[5].value][4].value + row[4].value

        else:
            recordlist = []
            name2recordlist[row[5].value] = recordlist
            titlerow = sheet.row(2) # 这一行是列名
            name2recordsumlist[row[5].value] = copy.deepcopy(row)
            recordlist.append(titlerow)
            recordlist.append(row)


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

def comp(row):
    if isinstance(row[4].value, str) :
        return 1999999999
    return row[4].value

for (k, v) in name2recordlist.items():
    persionname = k
    recordlist = v
    recordlist.sort(reverse=True, key=comp)

# 月度汇总表
recordlist = []
sumsheet = book_everymonth_everypersion.add_sheet("all", cell_overwrite_ok=True)
for (k, v) in name2recordsumlist.items():
    recordlist.append(v)

recordlist.sort(reverse=True, key=comp)
for rowi in range(len(recordlist)):
    row = recordlist[rowi]
    for colindex in range(len(row)):
        sumsheet.write(rowi, colindex, row[colindex].value)

    persionname = row[5].value
    everypersionRecordlist = name2recordlist[persionname]
    newsheet = book_everymonth_everypersion.add_sheet(persionname, cell_overwrite_ok=True)
    for rowi in range(len(everypersionRecordlist)):
        row = everypersionRecordlist[rowi]
        for colindex in range(len(row)):
            newsheet.write(rowi, colindex, row[colindex].value)


everymonth_everypersion_book_name = "book_everymonth_everypersion_" + month + ".xls"
book_everymonth_everypersion.save(everymonth_everypersion_book_name)



