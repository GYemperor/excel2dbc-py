# coding=utf-8

import xlrd
import re
import ctypes

str1 = "0x18EAFF00";
print(str(ctypes.c_uint32(eval(str1))))

# 打开文件
data = xlrd.open_workbook('pccH5.xlsx')

# 查看工作表
data.sheet_names()
print("sheets：" + str(data.sheet_names()))

# 通过文件名获得工作表,获取工作表1
table = data.sheet_by_name('CAN_Matrix')
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

rowvalue = table.row_values(1)
print(rowvalue)

fdbc = open("PCCH5.dbc","w+")

newContext="VERSION \"\"\n\n\nNS_ :\n\tNS_DESC_\n\
	CM_\n\
	BA_DEF_\n\
	BA_\n\
	VAL_\n\
	CAT_DEF_\n\
	CAT_\n\
	FILTER\n\
	BA_DEF_DEF_\n\
	EV_DATA_\n\
	ENVVAR_DATA_\n\
	SGTYPE_\n\
	SGTYPE_VAL_\n\
	BA_DEF_SGTYPE_\n\
	BA_SGTYPE_\n\
	SIG_TYPE_REF_\n\
	VAL_TABLE_\n\
	SIG_GROUP_\n\
	SIG_VALTYPE_\n\
	SIGTYPE_VALTYPE_\n\
	BO_TX_BU_\n\
	BA_DEF_REL_\n\
	BA_REL_\n\
	BA_DEF_DEF_REL_\n\
	BU_SG_REL_\n\
	BU_EV_REL_\n\
	BU_BO_REL_\n\
	SG_MUL_VAL_\n\nBS_:\n\nBU_:\n\n\n"
fdbc.write(newContext)

print("print noRow "+str(table.nrows))
noRow = 1
spaStr = " "
chID = ""
CintID = ""
intID = ""
while noRow < table.nrows:
    print("now row NO:"+str(noRow))
    noRowData = table.row_values(noRow)
    print("now row data:"+str(noRowData))
    print("rowdata[0]:"+noRowData[0])

    if noRowData[0] != "":
        chID = noRowData[2]
        print(chID)
        print(str(ctypes.c_uint32(eval(chID))))
        CintID = str(ctypes.c_uint32(eval(chID)))
        intID = re.sub("\D","",CintID)
        print("intID:"+str(intID))

        MsgName = noRowData[0]
        MsgDlc = noRowData[5]

        newContext = "BO_ " + str(intID) +spaStr + str(MsgName)+":" +spaStr+str(int(MsgDlc))+spaStr+"Vector__XXX\n"
        #print("newContext:\n"+newContext)
        #fdbc.write(newContext)
    if noRowData[0] == "":
        SignalName = noRowData[6]
        SignalDescribe = noRowData[7]
        SignalStartBit = noRowData[9]
        SignalBitLenth = noRowData[11]
        #SignalFactor = noRowData[13]
        SignalFactor = 0
        #SignalOffset = noRowData[14]
        SignalOffset = 0
        #SignalMin = noRowData[15]
        SignalMin = 0
        #SignalMax = noRowData[16]
        SignalMax = 0
        SignalUnit = noRowData[23]
        print("signal describe:\n"+str(re.sub("\W+","",SignalDescribe)))
        newContext = spaStr+"SG_" + spaStr+str(re.sub("\W+","",SignalName))+spaStr+":"+spaStr+str(int(SignalStartBit))+"|"+\
                 str(int(SignalBitLenth))+"@1+"+spaStr+"("+\
                str(float(SignalFactor))[0:-2]+","+str(float(SignalOffset))[0:-2]+")"+spaStr+"["+str(int(SignalMin))+"|"+str(int(SignalMax))+"]"+ \
                 spaStr+"\""+str(re.sub("\W+","",str(SignalUnit)))+"\""+spaStr+"Vector__XXX\n"
    print("total newContext:\n" + newContext)
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
