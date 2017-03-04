# -*- coding: utf-8 -*-

import xlsxwriter as wx
import sys

reload(sys)
sys.setdefaultencoding('utf8')

import xlrd


def open_excel(file='file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)

wb = wx.Workbook('hello.xlsx')

data = open_excel("hello.xlsx")
table = data.sheets()[0]
print table.cell(0, 1)

worksheet = wb.add_worksheet()
worksheet.write(0, 2, 'h677llo world')
worksheet.write('A1', 'H发大富翁')
# 向A1写入



# Some data we want to write to the worksheet.
expenses = (
    ['Rent', 1000],
    ['Gas', 100],
    ['Food', 300],
    ['Gym', 50],
)

# Start from the first cell. Rows and columns are zero indexed. 按标号写入是从0开始的，按绝对位置'A1'写入是从1开始的
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col, item)
    worksheet.write(row, col + 1, cost)
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SUM(B1:B4)')  # 调用excel的公式表达式

wb.close()
