# @Author  : Austin_Yee

import xlrd
import re

# 数据源以及多少门课程
source=r'C:\Users\17106\OneDrive - 汕头大学\课程表\1.0\课程表源.xls'
quantity =10

def getData(source,quantity):a'u's
    book = xlrd.open_workbook(source)
    s1 =book.sheet_by_index(0)

    course=[]

    for i in range (0,quantity):
        # 获取每一行数据
        classNumber=str(s1.cell(i,0).value)
        # 去除小数点和小数
        classNumber=classNumber[:-2]
        courseName=str(s1.cell(i,1).value)
        weekRange=str(s1.cell(i,4).value)
        whichclass1=str(s1.cell(i,5).value)
        whichclass2 =str(s1.cell(i,6).value)
        classroom=str(s1.cell(i,3).value)
        teacher=str(s1.cell(i,2).value)

        # 替换开课代码为班号
        courseName=re.sub(r'\[.*?\]','['+str(classNumber)+']',courseName)

        firstLine=courseName
        # 多节课的情况
        if whichclass2 == "":
            secondLine=weekRange+whichclass1
        else:
            secondLine=weekRange+whichclass1+','+whichclass2
        thirdLine=classroom
        fourthLine=teacher

        all=firstLine+'\n'+secondLine+'\n'+thirdLine+'\n'+fourthLine

        print('第'+str(i+1)+'行')
        print()
        print(all)
        print()

getData(source,quantity)