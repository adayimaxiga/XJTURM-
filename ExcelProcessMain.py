#! /usr/bin/env python
# -*- coding:utf-8 –*-

import xlrd
import xlwt
import string
#打开
Workdata = xlrd.open_workbook('员工刷卡记录表.xls')
SheetName = Workdata.sheet_names()
print(SheetName)        #打印名称
table = Workdata.sheets()[0]
nrows = table.nrows #行数
ncols = table.ncols #列数
#写入
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
SheetElectric = book.add_sheet('电控组', cell_overwrite_ok=True)
SheetElectric.write(0, 0, '姓名')
SheetElectric.write(0, 1, '时长')
ElecCount=0
SheetMechnical = book.add_sheet('机械组', cell_overwrite_ok=True)
SheetMechnical.write(0, 0, '姓名')
SheetMechnical.write(0, 1, '时长')
MeceCount=0
SheetVision = book.add_sheet('视觉组', cell_overwrite_ok=True)
SheetVision.write(0, 0, '姓名')
SheetVision.write(0, 1, '时长')
VersCount=0
SheetManagement = book.add_sheet('管理组', cell_overwrite_ok=True)
SheetManagement.write(0, 0, '姓名')
SheetManagement.write(0, 1, '时长')
ManaCount=0

def ProcessExcelRow(Rawlist):
    Processedlist = ['header']
    Flag = False
    for i in range(len(Rawlist)):
        if(Rawlist[i]!=''):
            Processedlist.append(Rawlist[i])
#            Processedlist.remove('header')
    return Processedlist
#声明一个类
class MessegeOfRobomasterPerson:
    def __init__(self,Name,Office,Number):
        self.name=Name
        self.office=Office
        self.number=Number

    def TimeSet(self,Time):
        self.time=round(Time,1)

    def print(self):
        print("姓名："+str(self.name))
        print("部门："+str(self.office))
        print("工号："+str(self.number))
        print("时长："+str(self.time))

    def ReturnName(self):
        return self.name

    def ReturnOffice(self):
        return self.office

    def ReturnNumber(self):
        return self.number

    def ReturnTime(self):
        return self.time

def safe_int(num):
    try:
        return int(num)
    except ValueError:
        result = []
        for c in num:
            if not ('0' <= c <= '9'):
                break
            result.append(c)
        if len(result) == 0:
            return 0
        return int(''.join(result))

def ConvertToHours(StringTime):
    BufferTimeHour = StringTime[0:2]
    BufferTimeMinute = StringTime[3:5]
    BufferTimeHourInt=safe_int(BufferTimeHour)
    BufferTimeMinuteInt=safe_int(BufferTimeMinute)
    return BufferTimeHourInt + BufferTimeMinuteInt/60

def isNum(value):
    try:
        value + 1
    except TypeError:
        return False
    else:
        return True

if __name__ == "__main__":
    #全局保存
    ZeroTimeCount = 0
    TimeRowsCount = 0
    TimePlus = 0
    FlagStartRecord = False

    for i in range(nrows):
        #循环初始变量
        FlagNameUpdate = False

        #获取原始行列
        ListTempDeleteTemp=ProcessExcelRow(table.row_values(i))
        #清楚无关行
        if(len(ListTempDeleteTemp)>1):
            ListTempDelete=ListTempDeleteTemp
        #打印
        print (ListTempDelete)
        for i in range(len(ListTempDelete)):
            if(ListTempDelete[i]=='姓名：'):
                Rname = ListTempDelete[i+1]
                ZeroTimeCount=0
                if(FlagStartRecord):
                    person.TimeSet(TimePlus)
                    TimePlus = 0
            elif(ListTempDelete[i]=='部门：'):
                Roffice = ListTempDelete[i+1]
                FlagNameUpdate = True
            elif(ListTempDelete[i]=='工号：'):
                Rnumber = ListTempDelete[i+1]
        if(FlagNameUpdate):
            if(FlagStartRecord):
                if(person.office=='管理组'):
                    ManaCount +=1
                    SheetManagement.write(ManaCount,0,person.ReturnName())
                    SheetManagement.write(ManaCount,1,person.ReturnTime())
                elif(person.ReturnOffice()=='电控组'):
                    ElecCount +=1
                    SheetElectric.write(ElecCount, 0, person.ReturnName())
                    SheetElectric.write(ElecCount, 1, person.ReturnTime())
                elif(person.ReturnOffice()=='机械组'):
                    MeceCount +=1
                    SheetMechnical.write(MeceCount, 0,  person.ReturnName())
                    SheetMechnical.write(MeceCount, 1,  person.ReturnTime())
                elif(person.ReturnOffice()=='视觉组'):
                    VersCount +=1
                    SheetVision.write(VersCount, 0, person.ReturnName())
                    SheetVision.write(VersCount, 1, person.ReturnTime())
                person.print()
            person = MessegeOfRobomasterPerson(Rname,Roffice,Rnumber)

        if(isNum(ListTempDelete[1])):
            ZeroTimeCount += 1
            FlagStartRecord = True
        else:
            if(ZeroTimeCount==1):
                for i in range(1,len(ListTempDelete)):
                    BufferTimeRawData=ListTempDelete[i]
                    if(BufferTimeRawData[7]==' '):
                        BufferFirstTimeValue=BufferTimeRawData[0:5]
                        TimeTempSave = ConvertToHours(BufferFirstTimeValue)
                        if (TimeTempSave<12):
                            TimePlus += 12 - TimeTempSave
                        elif(TimeTempSave<18):
                            TimePlus += 18 - TimeTempSave
                        elif(TimeTempSave<23):
                            TimePlus += 23 - TimeTempSave
                    else:
                        BufferFirstTimeValue=BufferTimeRawData[0:5]
                        BufferSecondTimeValue=BufferTimeRawData[6:11]
                        TimePlus += ConvertToHours(BufferSecondTimeValue) - ConvertToHours(BufferFirstTimeValue)
                TimeRowsCount +=1
            elif(ZeroTimeCount==2):
                Time=0
                TimeRowsCount=0
                TimePlus=0
                person.TimeSet(TimePlus)

book.save(r'RM打卡时间.xls')
