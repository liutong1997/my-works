#coding=utf-8
import docx


file=docx.Document(r'/home/liutong/桌面/my works/files_to_deal/3.docx')
tables = file.tables
for table in tables:
    if table.cell(0,0).text == '编号':
        m=0;n=0
        for i in range(len(table.rows)):
            m+=1#统计行数
            for j in range(len(table.rows[i].cells)):
                n+=1

    n=int(n/m)#统计列数
    for i in range(1,m):
        for j in range(n):
            for k in range(len(table.cell(i,j).paragraphs)):
                run = table.cell(i,j).paragraphs[k].clear()#清空每个单元格
            table.cell(1,0).merge(table.cell(m-1,n-1))#单元格合并
            
file.save(r'/home/liutong/桌面/my works/files_to_deal/3.docx')