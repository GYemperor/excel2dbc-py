# coding=utf-8

import xlrd
import xlwt
import re
import os

from xlutils.copy import copy
from xlwt import Style

# 打开文件
data = xlrd.open_workbook("808A6.xlsx")
wb = copy(data)
ws = wb.get_sheet(0)

#ws.write(0,0,"valu1",Style.default_style)
#wb.save('808_A-6.xls')

# 查看工作表
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

# colvalue = table.col_values(6)[2:]
# print(colvalue)
# i = 2
# for itemCol in colvalue:
#     if itemCol != "":
#         asplitstr = re.split("-", itemCol)
#         print(asplitstr)
#         bitsize = int(asplitstr[1]) - int(asplitstr[0]) + 1
#         print(bitsize)
#         ws.write(i, 8, str(bitsize), Style.default_style)
#     else:
#         ws.write(i, 8, "0", Style.default_style)
#     i += 1

colvalue = table.col_values(5)[2:]
print(colvalue)
i = 2
for itemCol in colvalue:
    if itemCol != "":
        asplitstr = re.split("te", itemCol)
        byteR = asplitstr[1]
        aspbytelist = re.split("-", byteR)
        print(aspbytelist)
        if table.col_values(6)[i] != "":
            aspbitsize = re.split("-", table.col_values(6)[i])
            startbitsize = int(aspbytelist[0])*8+int(aspbitsize[0])
            bitsize = int(aspbitsize[1]) - int(aspbitsize[0]) + 1
            print("start"+str(startbitsize))
        else:
            startbitsize = int(aspbytelist[0])*8
            print("aspbytelist")
            print(len(aspbytelist))
            if len(aspbytelist) == 2:
                bitsize = (int(aspbytelist[1]) - int(aspbytelist[0]) + 1) * 8
            if len(aspbytelist) == 1:
                bitsize = 8
        print("bitsize")
        print(bitsize)
        ws.write(i, 8, str(bitsize), Style.default_style)
        ws.write(i, 7, str(startbitsize), Style.default_style)

    i += 1
colvalue = table.col_values(9)[2:]
print(colvalue)
i = 2
for itemCol in colvalue:
    if itemCol.find("精度：") >= 0:
        asplitstr = re.split("度：", itemCol)
        print(asplitstr)

        if asplitstr[1].find("偏移量") >= 0:
            offsetR =  re.split("偏移量", asplitstr[1])
            print("pianyiliang")
            print(offsetR)
            if offsetR[-1].find("：") >= 0:
                offset = re.split("：",offsetR[-1])
                print(offset[-1])
                offs = offset[-1]
            else :
                print(offsetR[-1])
                offs = offsetR[-1]
            ws.write(i, 13, offs, Style.default_style)

            uintR = re.split("\d", offsetR[0].strip())
            print("uint")
            print(offsetR[0])
            print(offsetR[0].strip())
            print(uintR)
            print(uintR[-1])
            print(uintR[-1].strip())
            ws.write(i, 16, uintR[-1], Style.default_style)
            if offsetR[0].find("；") >= 0:
                print("this is ;")
                factorRw = re.split("；", offsetR[0].strip())
                print("factorRw:"+factorRw[0])
                factorP = re.split("；", uintR[-1].strip())
                print("factorP:"+ factorP[0])
                pstr = factorP[0]
            else :
                pstr = uintR[-1]
            try:
                factorR = re.split(pstr, factorRw[0])
                print("factor")
                factor = factorR[0].strip()
            except :
                print(factor)
            else:
                print(factor)
                ws.write(i, 12, factor, Style.default_style)
        else :
            factorR = re.split("\d", asplitstr[1])
            uintR = factorR[-1]
            ws.write(i, 16, uintR, Style.default_style)
            factorR = re.split(str(uintR), asplitstr[1])
            print("factor")
            print(factorR[0])
            ws.write(i, 12, factorR[0], Style.default_style)
    i += 1

wb.save('808A6.xls')