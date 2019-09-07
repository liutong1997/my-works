# coding=utf-8
from xlrd import *
from matplotlib.pyplot import *
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
import random
import numpy as np


# 随机颜色函数
def random_color():
    colorArr = ['1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']
    color = ""
    for i in range(6):
        color += colorArr[random.randint(0, 14)]
    return "#" + color


if __name__ == '__main__':

    fig = figure(figsize=(8, 8))
    ax1 = subplot(111, projection='3d')

    # 读取已知表格
    work_book = open_workbook(r'/home/liutong/桌面/my_works/files_to_deal/out1.xlsx')
    table = work_book.sheet_by_index(0)
    collector = []

    # 获取表格的数据
    for row in range(table.nrows):
        values = []
        for col in range(table.ncols):
            values.append(table.cell(row, col).value)
        collector.append(values)

    # 选择要对比的编号

    action = True
    list_choice = []
    count_num = 0
    none_list = []
    none_action = False

    while action:
        a_choice = input('请输入一个想要对比的序号，不得超过{},不得小于{},输入q或者Q表示退出选择:'.format(len(collector) - 1, 0))
        if a_choice == 'q' or a_choice == 'Q':

            # 非空检验
            for num in list_choice:
                count_num += 1
                if '' in collector[num]:
                    print('所选序号案列属性存在空值!以下是有空值的序号，请核对：')
                    none_list.append(num)
                    none_action = True
                else:
                    pass

            else:
                if none_action:
                    print(none_list)
                    count_num = 0
                    list_choice = []
                    none_action = False
                    none_list = []
                else:
                    action = False

        else:
            list_choice.append(eval(a_choice))

    # 数据对应字典

    name_dict = {'1': '内外向(E)', '2': '神经质(N)', '3': '精神质(P)', '4': '掩饰性(L)', '5': '躯体化', '6': '强迫症状',\
                 '7': '人际关系敏感', '8': '抑郁', '9': '焦虑', '10': '敌对', '11': '恐怖', '12': '偏执', \
                 '13': '精神病性', '14': '其他', '15': '总分', '16': '总均分', '17': '阳性项目数'}
    list_match = []

    # 选择对比数据
    action = True

    while action:
        a_match = input(
            "\n1.内外向(E),\n 2.神经质(N),\n3.精神质(P),\n 4.掩饰性(L),\n5.躯体化,\n 6.强迫症状,\n 7.人际关系敏感,\n 8.抑郁,\n 9.焦虑,\n 10.敌对,"
            "\n 11.恐怖,\n 12.偏执,\n 13.精神病性,\n 14.其他,\n 15.总分,\n 16.总均分,\n 17.阳性项目数\n请选择一个要对比的指标（输入q或者Q表示推出选择），输入序号就行：")
        if a_match == 'q' or a_match == 'Q':
            break
        else:
            list_match.append(a_match)

    # 获取数据

    m = 1
    for num in list_choice:
        name = collector[0]
        guy = collector[num]

        x_name = []
        y0 = []
        for n_ma in list_match:
            x_name.append(name_dict[n_ma])
            y0.append(guy[name.index(name_dict[n_ma])])
        y = list(map(float, y0))
        x = [n + 1 for n in range(len(x_name))]
        z = [m for n in range(len(list_match))]
        m += 1

        ax1.bar(x, y, z, zdir='y', alpha=0.6)

    # 规范坐标轴
    ax1.set_xlabel('属性名称', labelpad=26)
    ax1.set_ylabel('对比序号及性别', labelpad=18)
    ax1.set_zlabel('所选值')

    y_lname = list(map(int, list_choice))
    y_nname = []
    for n in y_lname:
        sex_index = name.index('性别')
        sex = collector[n][sex_index]
        y_nname.append('序号' + str(n) + '性别' + sex)
    y_lname = [x + 1 for x in range(len(list_choice))]
    yticks(y_lname, y_nname, rotation=-10)

    xticks(x, x_name, fontsize=8, rotation=50)
    title('自选人数自选属性对比柱状图')

    show()
    # fig.savefig('自选人数自选属性对比柱状图')
