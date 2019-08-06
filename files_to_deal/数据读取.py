#coding=utf-8
import docx
from xlsxwriter import *


workbook=Workbook('out1.xlsx')
worksheet = workbook.add_worksheet()#创建输出表格文件


def value_clear(dict2={'1':'a',"2":'s'}):
    key=dict2.keys()
    for i in key:
        dict2[i]=None
    return dict2
#定义清空字典值而不改变键的函数

def strs_are_nums(a='22.22'):
    n=0
    for i in range(len(a)):
        if '0'<=a[i]<='9':
            n+=1
        if a[i] == '.' and '0'<=a[i+1]<='9' and '0'<=a[i-1]<='9':
            n+=1
    if n == len(a):
        return True
    else:
        return False

#定义判断字符串是否只由浮点数或整数构成的函数

file=docx.Document(r'/home/liutong/桌面/my works/files_to_deal/3.docx' )
dict0={'序号':None,'年龄':None,'性别':None,'编号':None,'内外向(E)':None\
    ,'神经质(N)':None,'精神质(P)':None,'掩饰性(L)':None,\
    '躯体化':None,'强迫症状':None,'人际关系敏感':None,'抑郁':None,'焦虑':None,'敌对':None,'恐怖':None,\
    '偏执':None,'精神病性':None,'其他':None,'总分':None,'总均分':None,'阳性项目数':None}
#创建字典储存标题值

number={}
k=0;a=0;b=0#分别为用户数量，写入文件行数，列数

for key in dict0.keys():
    worksheet.write(a,b,key)
    b+=1
a+=1

#写入标题行
tables = file.tables
for table in tables:
    b=0
    if table.cell(0,1).text not in number and strs_are_nums(table.cell(0,1).text):#判断用户

        if k>1:
            for value in dict0.values():
                worksheet.write(a,b,value)
                b+=1
            a+=1
            #写入上一个用户信息

        value_clear(dict0)
        number[table.cell(0,1).text]=k
        k+=1
        #利用字典储存用户编号，并且统计用户数量

        for m in range(len(table.rows)):
            for n in range(len(table.rows[0].cells)):
                if table.cell(m,n).text in dict0:#判断键是否存在在字典中，再写入对应的值

                    if table.cell(m,n+1).text in ['男','女'] or strs_are_nums(table.cell(m,n+1).text):
                        dict0[table.cell(m,n).text] = table.cell(m,n+1).text
                        #写入值
                    try:
                        if strs_are_nums(table.cell(m,n+2).text):
                            dict0[table.cell(m,n).text] = table.cell(m,n+2).text
                            #利用try模块防止列表越界的情况，即n已经代表最后一列时，n+2是不存在的
                    except:
                        pass#直接跳过越界
    else:
        for m in range(len(table.rows)):
            for n in range(len(table.rows[0].cells)):
                if table.cell(m,n).text in dict0:
                    if table.cell(m,n+1).text in ['男','女'] or strs_are_nums(table.cell(m,n+1).text):
                        dict0[table.cell(m,n).text] = table.cell(m,n+1).text
                    try:
                        if strs_are_nums(table.cell(m,n+2).text):
                            dict0[table.cell(m,n).text] = table.cell(m,n+2).text                           
                    except:
                        pass

for value in dict0.values():
    worksheet.write(a,b,value)
    b+=1
a+=1
workbook.close()

