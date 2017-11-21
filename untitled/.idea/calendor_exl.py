# coding=UTF-8
import xlrd
import re
import types
from datetime import datetime
from xlrd import xldate_as_tuple

excel = xlrd.open_workbook('calendar_data.xlsx')

table = excel.sheet_by_index(1)

rows=table.nrows  #获取行数

cols=table.ncols    #获取列数
def unicodestr_encode(s):
    return '{'+','.join([hex(ord(c)) for c in s])+'},\n'

f = open('event_data.txt', 'w')
f.write('event_data = {\n')
for col in range(2, cols):
    cell_data = table.cell(0,col).value
    if type(cell_data) is types.UnicodeType:

        f.write(unicodestr_encode(cell_data.encode('utf-8')))
f.write('};')
f.flush()
f_time = open('event_time.txt', 'w')
f_time.write('time_data = {\n')
for row in range(1,rows):
    row_content = []
    f_time.write('{')
    for col in range(0,cols):
        if col ==1 :
            continue
        ctype = table.cell(row,col).ctype  # 表格的数据类型
        cell_data = table.cell(row, col).value
        if ctype == 2 and cell_data % 1 == 0:  # 如果是整形
            cell_data = int(cell_data)
        elif ctype == 3:
            # 转成datetime对象
            date = datetime(*xldate_as_tuple(cell_data, 0))
            if col==0 :
                cell_data = date.strftime('%Y,%m,%d,{')
            else :
                cell_data = date.strftime('\"%H:%M%p\",')
        elif ctype == 4:
            cell_data = True if cell_data == 1 else False
        row_content.append(cell_data)
        f_time.write(cell_data)
    #f_time.write(','.join([row_content]))
    f_time.write('}},\n')

f_time.write('};')
f_time.flush()
