import os
import sys

import xlrd
import xlwt
import copy

name_colindex = 5  # “收款单位”  所在的列 index
colname_rowindex = 2  # 列名所在的index
weight_colindex = 2  # "重量" 所在的列index
hanshuie_colindex = 3  # "含税额" 列index
fukuanjiane_colindex = 4  # "付款金额" 列index

summary_colindex_list = [2, 3, 4]

def getColChar(colindex):
    return chr(colindex + 65) + "列"

def getRowNumber(rowindex):
    return str(rowindex + 1) + "行"

class MonthBook:

    def getSummarys(self):
        return self.name2recordsum

    def __init__(self, excelname):
        self.excelname = excelname
        self.name2recordlist = {}  # 每个人一个月内的记录
        self.name2recordsum = {}  # 一个月的汇总
        beginIndex = excelname.find("_", 0, len(excelname)) + 1
        endIndex = excelname.rfind(".", 0, len(excelname))
        self.month = excelname[beginIndex:endIndex]

        # print(self.month)

        self.excel_fullname = os.getcwd() + '\\' + excelname
        self.rdata = xlrd.open_workbook(self.excel_fullname)
        # print('sheets nums:', rdata.nsheets)  # excel sheets 个数
        self.titlerow = None

        # 汇总每个月的每个人
        for sheet in self.rdata.sheets():  # 每个月内的每一天是一个sheet
            # print("open sheet name:", sheet.name)
            for rowindex in range(sheet.nrows):
                if rowindex <= colname_rowindex:  # 前两行是公司名
                    if self.titlerow == None and rowindex == colname_rowindex:
                        self.titlerow = sheet.row(rowindex)
                    continue
                row = sheet.row(rowindex)
                if row[name_colindex].value == "":
                    continue
                # 以下代码用来检查该字段能不能汇总, 仅仅用于错误检查
                for colindex in summary_colindex_list:
                    try:
                        testfieldtype = row[colindex].value + 1
                    except TypeError:
                        print("表 " + self.excelname + " sheet:" + sheet.name +" 单元格:" + getColChar(colindex) + getRowNumber(
                            rowindex) + "不支持加法，请检查是否有数字,", TypeError)

                if row[name_colindex].value in self.name2recordlist:
                    recordlist = self.name2recordlist[row[name_colindex].value]
                    recordlist.append(row)
                else:
                    recordlist = []
                    self.name2recordlist[row[name_colindex].value] = recordlist
                    self.titlerow = sheet.row(colname_rowindex)  # 这一行是列名
                    recordlist.append(self.titlerow)
                    recordlist.append(row)

        for (k, v) in self.name2recordlist.items():
            persionname = k
            recordlist = v
            recordlist.sort(reverse=True, key=comp)

    def summary(self):
        # 每个人按月汇总
        book_month_summary = xlwt.Workbook(encoding='utf-8')
        # 月度汇总表
        sumsheet = book_month_summary.add_sheet("all", cell_overwrite_ok=True)
        sumrecordlist = []
        for (k, v) in self.name2recordlist.items():
            i = 0
            for record in v:
                i = i + 1
                if i == 1:
                    continue  # 列名不要加进去了
                if k in self.name2recordsum:
                        self.name2recordsum[k][weight_colindex].value = self.name2recordsum[k][weight_colindex].value + \
                                                                        record[weight_colindex].value
                        self.name2recordsum[k][hanshuie_colindex].value = self.name2recordsum[k][hanshuie_colindex].value + \
                                                                          record[hanshuie_colindex].value
                        self.name2recordsum[k][fukuanjiane_colindex].value = self.name2recordsum[k][
                                                                                 fukuanjiane_colindex].value + record[
                                                                                 fukuanjiane_colindex].value
                else:
                    recordcopy = copy.deepcopy(record)
                    self.name2recordsum[k] = recordcopy
                    sumrecordlist.append(recordcopy)

        sumrecordlist.append(self.titlerow)
        sumrecordlist.sort(reverse=True, key=comp)
        for rowi in range(len(sumrecordlist)):
            row = sumrecordlist[rowi]
            if row != self.titlerow:
                row[0].value = rowi
            for colindex in range(len(row)):
                if colindex <= name_colindex:
                    sumsheet.write(rowi, colindex, row[colindex].value)
            if row == self.titlerow:
                continue
            persionname = row[name_colindex].value
            everypersionRecordlist = self.name2recordlist[persionname]
            newsheet = book_month_summary.add_sheet(persionname, cell_overwrite_ok=True)
            for rowi in range(len(everypersionRecordlist)):
                row = everypersionRecordlist[rowi]
                if row != self.titlerow:
                    row[0].value = rowi
                for colindex in range(len(row)):
                    if colindex <= name_colindex:
                        newsheet.write(rowi, colindex, row[colindex].value)

        month_summary_name = "./output/book_month_summary_" + self.month + ".xls"
        book_month_summary.save(month_summary_name)
        print("     生成success：", month_summary_name)


def comp(row):
    if isinstance(row[fukuanjiane_colindex].value, str):
        return sys.maxsize
    return row[fukuanjiane_colindex].value


class YearStatistics:
    monthbooks = []
    name2recordsumlist = {}

    def __init__(self):
        pass

    def addMonthBook(self, monthbook):
        YearStatistics.monthbooks.append(monthbook)

    def summary(self):
        # 每个人按月汇总
        book_year_summary = xlwt.Workbook(encoding='utf-8')
        # 月度汇总表
        sumsheet = book_year_summary.add_sheet("all", cell_overwrite_ok=True)
        name2summary = {}
        summarylist = []
        titleRow = None
        for book in YearStatistics.monthbooks:
            monthsummarys = book.getSummarys()
            if titleRow == None:
                titleRow = book.titlerow
            for (name, summary) in monthsummarys.items():
                if name in name2summary:
                    name2summary[name][weight_colindex].value = name2summary[name][weight_colindex].value + summary[
                        weight_colindex].value
                    name2summary[name][hanshuie_colindex].value = name2summary[name][hanshuie_colindex].value + summary[
                        hanshuie_colindex].value
                    name2summary[name][fukuanjiane_colindex].value = name2summary[name][fukuanjiane_colindex].value + \
                                                                     summary[fukuanjiane_colindex].value
                else:
                    summaryCopy = copy.deepcopy(summary)
                    name2summary[name] = summaryCopy
                    summarylist.append(summaryCopy)

        summarylist.append(titleRow)
        summarylist.sort(reverse=True, key=comp)
        for rowi in range(len(summarylist)):
            row = summarylist[rowi]
            if row != titleRow:
                row[0].value = rowi
            for colindex in range(len(row)):
                if colindex <= name_colindex:
                    sumsheet.write(rowi, colindex, row[colindex].value)

        year_summary_name = "./output/book_year_summary" + ".xls"
        book_year_summary.save(year_summary_name)
        print("     生成success：", year_summary_name)


def main():
    yearStatictics = YearStatistics()

    for root, dirs, files in os.walk("./input", topdown=False):
        for name in files:
            filename = os.path.join(root, name)
            print("读取：", os.path.join(root, name))
            monthbook = MonthBook(filename)
            monthbook.summary()

            yearStatictics.addMonthBook(monthbook)

    yearStatictics.summary()

    os.system('pause')


if __name__ == '__main__':
    main()


def test():
    pass