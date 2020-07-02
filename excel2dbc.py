# coding=utf-8

import xlrd
import re

# 打开文件
data = xlrd.open_workbook('can.xlsx')

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

rowvalue = table.row_values(0)
print(rowvalue)

fdbc = open("dbcconfig.dbc","w+")
newContext = "VERSION \"\" \n"
fdbc.write(newContext)

newContext = "NS_ : \n"
fdbc.write(newContext)
listNS = {
    'NS_DESC_',
    'CM_',
    'BA_DEF_',
    'BA_',
    'VAL_',
    'CAT_DEF_',
    'CAT_',
    'FILTER',
    'BA_DEF_DEF_',
    'EV_DATA_',
    'ENVVAR_DATA_',
    'SGTYPE_',
    'SGTYPE_VAL_',
    'BA_DEF_SGTYPE_',
    'BA_SGTYPE_',
    'SIG_TYPE_REF_',
    'VAL_TABLE_',
    'SIG_GROUP_',
    'SIG_VALTYPE_',
    'SIGTYPE_VALTYPE_',
    'BO_TX_BU_',
    'BA_DEF_REL_',
    'BA_REL_',
    'BA_DEF_DEF_REL_',
    'BU_SG_REL_',
    'BU_EV_REL_',
    'BU_BO_REL_',
    'SG_MUL_VAL_',}
print(listNS)
for itemNS in listNS:
    print(itemNS)
    newContext = "\t"+itemNS+"\n"
    fdbc.write(newContext)

newContext = "BS_ : \n"
fdbc.write(newContext)
newContext="\n"
fdbc.write(newContext)
newContext = "BU_ : \n"
fdbc.write(newContext)
newContext="\n"
fdbc.write(newContext)

print("print noRow "+str(table.nrows))
noRow = 1
spaStr = " "
chID = ""
intID = 0
while noRow < table.nrows:
    print(str(noRow))
    noRowData = table.row_values(noRow)
    print(noRowData)
    print(noRowData[0])
    print(chID)

    if noRowData[0] != chID:
        chID = noRowData[0]
        print(chID)
        intID = int(chID,16)
        print(str(intID))

        newContext = "BO_ " + str(intID) +spaStr + "_"+chID+":" +spaStr+str(int(noRowData[1]))+spaStr+"Vector__XXX\n"
        print("newContext"+newContext)
        fdbc.write(newContext)
    newContext = spaStr+"SG_" + spaStr+str(noRowData[2])+spaStr+":"+spaStr+str(int(noRowData[3]))+"|"+\
                 str(int(noRowData[4]))+"@"+str(int(noRowData[5]))+str(noRowData[6])+spaStr+"("+\
                str(noRowData[7])+","+str(noRowData[8])+")"+spaStr+"["+str(noRowData[9])+"|"+str(noRowData[10])+"]"+ \
                 spaStr+"\""+str(noRowData[11])+"\""+spaStr+"Vector__XXX\n"
    print("newContext" + newContext)
    fdbc.write(newContext)
    noRow+=1

newContext="\n"
fdbc.write(newContext)
newContext = "BA_DEF_  \"BusType\" STRING ;\n\
BA_DEF_ BU_  \"NodeLayerModules\" STRING ;\n\
BA_DEF_ BU_  \"ECU\" STRING ;\n\
BA_DEF_ BU_  \"CANoeJitterMax\" INT 0 0;\n\
BA_DEF_ BU_  \"CANoeJitterMin\" INT 0 0;\n\
BA_DEF_ BU_  \"CANoeDrift\" INT 0 0;\n\
BA_DEF_ BU_  \"CANoeStartDelay\" INT 0 0;\n\
BA_DEF_DEF_  \"BusType\" \"\";\n\
BA_DEF_DEF_  \"NodeLayerModules\" \"\";\n\
BA_DEF_DEF_  \"ECU\" \"\";\n\
BA_DEF_DEF_  \"CANoeJitterMax\" 0;\n\
BA_DEF_DEF_  \"CANoeJitterMin\" 0;\n\
BA_DEF_DEF_  \"CANoeDrift\" 0;\n\
BA_DEF_DEF_  \"CANoeStartDelay\" 0;\n\
BA_ \"BusType\" \"CAN\";\n"
print("newContext"+newContext)
fdbc.write(newContext)
