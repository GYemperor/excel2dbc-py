# coding=utf-8

import xlrd
import xlwt
import re

# 打开文件
data = xlrd.open_workbook('can651.xlsx')

# 查看工作表
data.sheet_names()
print("sheets：" + str(data.sheet_names()))

# 通过文件名获得工作表,获取工作表1
table = data.sheet_by_name('Sheet1')
# 获取行数和列数
# 行数：table.nrows
# 列数：table.ncols
print("总行数：" + str(table.nrows))
print("总列数：" + str(table.ncols))

# 获取整行的值 和整列的值，返回的结果为数组
# 整行值：table.row_values(start,end)
# 整列值：table.col_values(start,end)
# 参数 start 为从第几个开始打印，
# end为打印到那个位置结束，默认为none
#print("整行值：" + str(table.row_values(0)))
#print("整列值：" + str(table.col_values(9)))

colvalue = table.col_values(9)[1:]
print(colvalue)

asplitstr = re.split(" ",colvalue[0])
asplit = asplitstr[:-1]
print( asplitstr)
print( asplit)

newLine = 0

wb = xlwt.Workbook()
# 添加sheet
ws = wb.add_sheet('data', cell_overwrite_ok=True)
# 注意这里的index和后面的i，不要混淆
ws.write(0, 0, "DID")
ws.write(0,1,"LEN")


i = 1
for colitem in colvalue:
    #print(colitem)
    asplitstr = re.split(" ", colitem)
    asplit = asplitstr[:-1]
    if asplit[1]=='62' or asplit[2]=='62':
        if asplit[1]=='62':
            ws.write(i, 0, asplit[0])
            ws.write(i, 1, asplit[2]+' '+asplit[3])
        if asplit[2]=='62':
            ws.write(i, 0, asplit[1])
            ws.write(i, 1, asplit[3] +' '+ asplit[4])
        i+=1
wb.save("did.xls")




