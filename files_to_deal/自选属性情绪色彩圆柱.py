#coding=utf-8
from xlrd import *
from matplotlib.pyplot import *
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
import random
import numpy as np

#随机颜色函数
def randomcolor():
    colorArr = ['1','2','3','4','5','6','7','8','9','A','B','C','D','E','F']
    color = ""
    for i in range(6):
        color += colorArr[random.randint(0,14)]
    return "#"+color

fig = figure(figsize=(8,8))
ax1 = subplot(111,projection='3d')

#读取已知表格
work_book = open_workbook(r'/home/liutong/桌面/my works/files_to_deal/out1.xlsx'  )
table = work_book.sheet_by_index(0)
collector = []

#获取表格的数据
for row in range(table.nrows):
    values = []
    for col in range(table.ncols):
        values.append(table.cell(row,col).value)
    collector.append(values)

#选择要对比的编号
action = True

while action:
    a_choice = input('请输入一个想要对比的序号，不得超过{},不得小于{}:'.format(len(collector)-1,0)) 
    num = eval(a_choice)
    if num > 0 and num <= len(collector)-1:
        guy = collector[num]
        if '' in guy:
            print('注意！！！所选序号属性存在空值，请重新选择\n')
            print('请检查原病人信息数据表是否有信息缺失！！！\n')
            continue
        else:
            break
    else:
        print('输入的值不规范，shuru请重新输入有效值')
        

#数据对应字典
name_dict = {'1':'内外向(E)', '2':'神经质(N)','3':'精神质(P)', '4':'掩饰性(L)',\
        '5':'躯体化', '6':'强迫症状', '7':'人际关系敏感', '8':'抑郁', '9':'焦虑', '10':'敌对', '11':'恐怖', '12':'偏执', '13':'精神病性', '14':'其他', '15':'总分', '16':'总均分', '17':'阳性项目数'}
list_match = []

#选择对比数据
while  action:
    a_match = input("\n1.内外向(E),\n 2.神经质(N),\n3.精神质(P),\n 4.掩饰性(L),\n5.躯体化,\n 6.强迫症状,\n 7.人际关系敏感,\n 8.抑郁,\n 9.焦虑,\n 10.敌对,\n 11.恐怖,\n 12.偏执,\n 13.精神病性,\n 14.其他,\n 15.总分,\n 16.总均分,\n 17.阳性项目数\n请选择一个要对比的指标（输入q或者Q表示推出选择），输入序号就行：")
    if a_match == 'q' or a_match == 'Q':
        break
    else:
        list_match.append(a_match)

#获取数据

name = collector[0]
guy = collector[num]
z_name = []   

y0 = []
for n_ma in list_match:
    z_name.append(name_dict[n_ma])
    y0.append(guy[name.index(name_dict[n_ma])])
r_list = list(map(float,y0))

#画坐标轴
k = int(max(r_list) + 2)
tic = np.linspace(-k,k,10)
xticks(tic)
ztic = [n+1 for n in range(len(z_name))]
ax1.set_zticks(ztic)
ax1.set_zticklabels(z_name,fontsize=9)
xlabel('x轴')
ylabel('y轴')
ax1.set_zlabel('人格因子',labelpad=12)

#画圆柱
bottom = 0
top = 1
for n in range(len(r_list)):
    theta = np.linspace(0,2*np.pi,50)
    h = np.linspace(bottom,top,2)
    y = np.outer(r_list[n]*np.sin(theta),np.ones(len(h)))
    x = np.outer(r_list[n]*np.cos(theta),np.ones(len(h)))
    z = np.outer(np.ones(len(theta)),h)
    color1 = randomcolor()
    ax1.plot_surface(x,y,z,alpha=0.5,color = color1)
    ax1.text(r_list[n]*np.cos(2*np.pi),r_list[n]*np.sin(2*np.pi),(bottom+top)/2,z_name[n],fontsize=12)
    bottom += 1
    top += 1
title('值越大半径越大的情绪圆柱图')


show()
fig.savefig('序号'+str(num)+'自选属性的情绪圆柱图')