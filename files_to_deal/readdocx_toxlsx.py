# coding=utf-8
import shutil

import docx
from xlsxwriter import *
from multiprocessing import Process
import os
import glob


# 定义自动创建文件夹的函数
def make_folder(p):
    if os.path.exists(p):  # 判断文件夹是否存在
        shutil.rmtree(p)  # 如果存在删除原有目录的文件夹以及其中所有文件
    os.mkdir(p)


def collect_information(t):
    if t.cell(0, 0).text == '编号':  # 判断表格第一个单元格是否为编号
        if int(t.cell(0, 3).text) == len(list0):  # 判断序号是否与已经读取的用户长度一致
            # 变量a储存序号
            global a
            a = int(t.cell(0, 3).text)
            list0.append(dict0.copy())  # 列表list0储存每个用户的信息字典
            for m in range(1, len(t.rows)):
                for n in range(0, len(t.rows[0].cells)):
                    if t.cell(m, n).text in dict0.keys():
                        list0[int(t.cell(0, 3).text)][t.cell(m, n).text] = t.cell(m, n + 3).text
                    if t.cell(m, n).text == '身份证号':
                        list0[int(t.cell(0, 3).text)][t.cell(m, n).text] = t.cell(m, n + 3).text[0:6]
        # 遍历首个单元格为编号的表的首行数值
        for n in range(0, len(t.rows[0].cells)):
            if t.cell(0, n).text in dict0.keys():
                list0[int(t.cell(0, 3).text)][t.cell(0, n).text] = t.cell(0, n + 1).text  # 将列表中字典的对应值改为读取到的数值
    if t.cell(0, 0).text == '因子名称' and t.cell(1, 0).text in ['躯体化']:  # 读取首个单元格是因子名称的表
        # 遍历该表所有行
        for m in range(1, len(t.rows)):
            list0[a][t.cell(m, 0).text] = t.cell(m, 2).text  # 将列表中字典的对应值改为读取到的数值
    if t.cell(0, 0).text == '因子名称' and t.cell(1, 0).text in ['生命质量']:  # 读取首个单元格是因子名称的表
        # 遍历该表所有行
        for m in range(1, len(t.rows)):
            list0[a][t.cell(m, 0).text] = t.cell(m, 2).text  # 将列表中字典的对应值改为读取到的数值
    if t.cell(0, 0).text == '因子名称' and t.cell(1, 0).text in ['内外向(E)']:
        for m in range(1, len(t.rows)):
            list0[a][t.cell(m, 0).text + t.cell(0, 2).text] = t.cell(m, 2).text  # 将列表中字典的对应值改为读取到的数值
            list0[a][t.cell(m, 0).text + t.cell(0, 3).text] = t.cell(m, 3).text  # 将列表中字典的对应值改为读取到的数值
    if t.cell(0, 0).text == '测验工具' and t.cell(1, 0).text == '测验日期':
        for m in range(0, len(t.rows)):
            list0[a][t.cell(m, 0).text] = t.cell(m, 1).text
            list0[a][t.cell(m, 2).text] = t.cell(m, 3).text


def wt_xlsx(path):
    work_book = Workbook(path + '/new.xlsx')  # 创建／添加excel工作簿

    worksheet = work_book.add_worksheet()  # 创建输出表格文件
    # 创建a、b变量存放行列值
    a = 0;
    b = 0
    # 写入标题行
    for key in dict0.keys():
        worksheet.write(a, b, key)  # 写入第一行属性值
        b += 1  # 列自增
    b = 0  # 列重置
    a += 1  # 行自增
    # 遍历储存所有字典的列表，取出每个用户信息的字典
    for n in range(1, len(list0)):
        # 取出字典中每个值的信息
        for value in list0[n].values():
            worksheet.write(a, b, value)  # 每个空写入对应的值
            b += 1  # 列自增
        b = 0  # 列重置
        a += 1  # 行自增
    work_book.close()  # 关闭工作簿


if __name__ == '__main__':
    # 创建字典储存标题值
    dict0 = {'序号': None, '年龄': None, '性别': None, '编号': None, '学校': None, '专业': None, '住址': None, '身份证号': None, \
             '入学年份': None, '备注': None, '院系': None, '内外向(E)原始分': None, '内外向(E)标准分': None, '测验工具': None, \
             '测验用时': None, '神经质(N)原始分': None, '神经质(N)标准分': None, '精神质(P)原始分': None, '精神质(P)标准分': None, \
             '掩饰性(L)原始分': None, '掩饰性(L)标准分': None, '躯体化': None, '强迫症状': None, '人际关系敏感': None, '抑郁': None, \
             '焦虑': None, '敌对': None, '恐怖': None, '偏执': None, '精神病性': None, '其他': None, '总分': None, '总均分': None,
             '阳性项目数': None, '生命质量': None, '曾有自杀行为': None, '自杀意念': None, '近一年觉得活着没意思': None, \
             '近一年觉得死了才好': None, '近一年有自杀念头': None, '近一月有自杀念头': None}
    list0 = [dict0]  # 创建列表list0储存
    base_path = os.getcwd()
    file_path = base_path + '/qingxixinxi'
    save_path = base_path + '/duqushuju'
    make_folder(save_path)
    for file_name in glob.glob(os.path.join(file_path, '*.docx')):
        file = docx.Document(file_name)
        tables = file.tables  # 读取所有表格
        # 遍历所有表格
        for n in range(len(tables)):
            p = Process(collect_information(tables[n]))  # 调用收集用户信息的函数
            # 启动一个子进程来运行子任务
            p.start()
            p.join()
            # 子进程完成后，继续运行主进程
        # 调用函数写入xlsx文件
    wt_xlsx(save_path)
