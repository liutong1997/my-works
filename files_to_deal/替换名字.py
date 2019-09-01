#coding=utf-8
import docx


file=docx.Document(r'/home/liutong/桌面/my works/files_to_deal/1.docx' )
tables = file.tables
name={}#创建字典用于储存替换名字和编号，防止同名同姓
a=1#创建替换名字的编号

for table in tables:
    if table.cell(0,0).text == '编号':#查找目标表

       if table.cell(0,3).text+table.cell(0,1).text not in name:#修改目标表的姓名属性
           name[table.cell(0,3).text+table.cell(0,1).text]=a#用字典储存编号
           a+=1
           b=str(name.get(table.cell(0,3).text+table.cell(0,1).text))#将编号转换为字符串
           table.cell(0,3).paragraphs[0].clear()
           table.cell(0,3).paragraphs[0].add_run(b)#用编号替换名字
           c='序号'
           table.cell(0,2).paragraphs[0].clear()
           table.cell(0,2).paragraphs[0].add_run(c)

       else:
           b=str(name.get(table.cell(0,3).text+table.cell(0,1).text))#将编号转换为字符串
           table.cell(0,3).paragraphs[0].clear()
           table.cell(0,3).paragraphs[0].add_run(b)#用编号替换名字
           c='序号'
           table.cell(0,2).paragraphs[0].clear()
           table.cell(0,2).paragraphs[0].add_run(c)

# print(name)

file.save(r'/home/liutong/桌面/my works/files_to_deal/3.docx' )