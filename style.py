import xlwt


def set_style(name, height, bold=True):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name
    font.bold = bold
    # font.color_index = 4
    font.height = height

    style.font = font
    return style


#         # style = xlwt.XFStyle()  # 初始化样式
#         # font = xlwt.Font()  # 为样式创建字体
#         # font.name = 'Times New Roman'
#         # font.bold = True  # 黑体
#         # font.underline = True  # 下划线
#         # font.italic = True  # 斜体字
#         # style.font = font  # 设定样式
#         #
#         # borders = xlwt.Borders()
#         # borders.left = xlwt.Borders.THIN
#         # borders.left = xlwt.Borders.THIN
#         # # NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
#         # # THIN： 官方代码中THIN所表示的值为1，边框为实线
#         # # 左边框 细线
#         # borders.left = 1
#         # # 右边框 中细线
#         # borders.right = 2
#         # # 上边框 虚线
#         # borders.top = 3
#         # # 下边框 点线
#         # borders.bottom = 4
#         # # 内边框 粗线
#         # borders.diag = 5
#         #
#         # # 左边框颜色 蓝色
#         # borders.left_colour = 0x0C
#         # # 右边框颜色 金色
#         # borders.right_colour = 0x33
#         # # 上边框颜色 绿色
#         # borders.top_colour = 0x11
#         # # 下边框颜色 红色
#         # borders.bottom_colour = 0x0A
#         # # 内边框 黄色
#         # borders.diag_colour = 0x0D
#         # # 定义格式
#         # style.borders = borders
#
#

################################