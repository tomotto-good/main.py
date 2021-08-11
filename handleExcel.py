from xlrd import open_workbook
from xlutils.copy import copy
import xlwt


class handleExcel:
    def __init__(self, excelName):
        workBook = xlwt.Workbook(encoding='utf-8')
        pattern = xlwt.Pattern()  # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5
        # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        self.style = xlwt.XFStyle()  # Create the Pattern
        self.style.pattern = pattern  # Add Pattern to Style
        workBook.add_sheet('用户操作日志')
        workBook.add_sheet('统计')
        workBook.save(excelName)

    def writeSheet1(self, excelName, vaules):
        rexcel = open_workbook(excelName)
        excel = copy(rexcel)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
        table = excel.get_sheet(0)  # 用xlwt对象的方法获得要操作的sheet
        table.write(0, 0, '用户名', self.style)
        table.write(0, 1, '手机号', self.style)
        table.write(0, 2, '用户操作', self.style)
        table.write(0, 3, '操作时间', self.style)
        table.write(0, 4, '登录类型', self.style)
        for i, item in enumerate(vaules):
            table.write(i + 1, 0, item[0])
            table.write(i + 1, 1, item[1])
            table.write(i + 1, 2, item[2])
            table.write(i + 1, 3, item[3])
            table.write(i + 1, 4, item[4])
        excel.save(excelName)

    def writeSheet2(self, excelName, vaules):
        rexcel = open_workbook(excelName)
        excel = copy(rexcel)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
        table = excel.get_sheet(1)  # 用xlwt对象的方法获得要操作的sheet
        table.write(0, 0, '用户名', self.style)
        table.write(0, 1, '手机号', self.style)
        table.write(0, 2, '操作次数', self.style)
        table.write(0, 3, 'web', self.style)
        table.write(0, 4, 'android', self.style)
        table.write(0, 5, 'ios', self.style)
        table.write(0, 6, 'h5', self.style)
        table.write(0, 7, '备注', self.style)
        table.write(0, 8, '所有企业', self.style)
        table.write(0, 9, '当前企业', self.style)
        for i, item in enumerate(vaules):
            table.write(i + 1, 0, item[0])
            table.write(i + 1, 1, item[1])
            table.write(i + 1, 2, item[2])
            table.write(i + 1, 3, item[3])
            table.write(i + 1, 4, item[4])
            table.write(i + 1, 5, item[5])
            table.write(i + 1, 6, item[6])
            table.write(i + 1, 7, item[7])
            table.write(i + 1, 8, item[8])
            table.write(i + 1, 9, item[9])
        excel.save(excelName)
