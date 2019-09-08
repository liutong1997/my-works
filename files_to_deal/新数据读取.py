# coding=utf-8
import docx
from xlsxwriter import *


def collect_information(t):
    # 判断表格第一个单元格是否为编号
    if t.cell(0, 0).text == '编号':
        # 判断序号是否与已经读取的用户长度一致
        if int(t.cell(0, 3).text) == len(list0):
            # 变量a储存序号
            a = int(t.cell(0, 3).text)
            # 列表list0储存每个用户的信息字典
            list0.append(dict0.copy())
        # 遍历首个单元格为编号的表的首行数值
        for n in range(1, len(t.rows[0].cells), 2):
            # 将列表中字典的对应值改为读取到的数值
            list0[a][t.cell(0, n - 1).text] = t.cell(0, n).text
    # 读取首个单元格是因子名称的表
    if t.cell(0, 0).text == '因子名称' and t.cell(1, 0).text in ['内外向(E)', '躯体化']:
        # 遍历该表所有行
        for m in range(1, len(t.rows)):
            # 将列表中字典的对应值改为读取到的数值
            list0[a][t.cell(m, 0).text] = t.cell(m, 2).text


def wt_xlsx():
    # 创建／添加excel工作簿
    work_book = Workbook(input('请输入正确的输出表格文件路径(没有自动创建)，保存格式建议为较稳定的xlsx：'))
    # 创建输出表格文件
    worksheet = workbook.add_worksheet()
    # 创建a、b变量存放行列值
    a = 0;
    b = 0
    # 写入标题行
    for key in dict0.keys():
        # 写入第一行属性值
        worksheet.write(a, b, key)
        # 列自增
        b += 1
    # 列重置
    b = 0
    # 行自增
    a += 1
    # 遍历储存所有字典的列表，取出每个用户信息的字典
    for n in range(1, len(list0)):
        # 取出字典中每个值的信息
        for value in list0[n].values():
            # 每个空写入对应的值
            worksheet.write(a, b, value)
            # 列自增
            b += 1
        # 列重置
        b = 0
        # 行自增
        a += 1
    # 关闭工作簿
    work_book.close()


if __name__ == '__main__':
    # 输入正确的文件路径
    file = docx.Document(input('请输入要读取的正确的docx文件路径:'))
    # 读取所有表格
    tables = file.tables
    # 创建字典储存标题值
    dict0 = {'序号': None, '年龄': None, '性别': None, '编号': None, '内外向(E)': None
        , '神经质(N)': None, '精神质(P)': None, '掩饰性(L)': None, \
             '躯体化': None, '强迫症状': None, '人际关系敏感': None, '抑郁': None, '焦虑': None, '敌对': None, '恐怖': None, \
             '偏执': None, '精神病性': None, '其他': None, '总分': None, '总均分': None, '阳性项目数': None}
    # 创建列表list0储存
    list0 = [dict0]
    # 遍历所有表格
    for table in tables:
        # 调用收集用户信息的函数
        collect_information(table)
    # 调用函数写入xlsx文件
    wt_xlsx()
