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

#3d柱状图函数
def bar_of_3d(x1,x2,x3,x4,y1,y2,y3,y4,h,color3D):
    

#3d圆柱体底面
    x = np.array([[x1,x2],[x3,x4]])
    y = np.array([[y1,y2],[y3,y4]])
    z = np.zeros((2,2))
    ax1.plot_surface(x,y,z,color = color3D)

#3d柱体侧面
    x = np.array([[x1,x3],[x1,x3]])
    y = y1*np.ones((2,2))
    z = np.array([[0,0],[h,h]])
    ax1.plot_surface(x,y,z,color = color3D)

    x = np.array([[x2,x4],[x2,x4]])
    y = y2*np.ones((2,2))
    z = np.array([[0,0],[h,h]])
    ax1.plot_surface(x,y,z,color = color3D)

    x = np.array([[x1,x2],[x1,x2]])
    z = np.array([[0,0],[h,h]])
    y = np.array([[y1,y2],[y1,y2]])
    ax1.plot_surface(x,y,z,color = color3D)

    x = np.array([[x3,x4],[x3,x4]])
    z = np.array([[0,0],[h,h]])
    y = np.array([[y1,y2],[y1,y2]])
    ax1.plot_surface(x,y,z,color = color3D)

#柱顶部
    x = np.array([[x1,x2],[x3,x4]])
    y = np.array([[y1,y2],[y3,y4]])
    z = h*np.ones((2,2))
    ax1.plot_surface(x,y,z,color = color3D)



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
list_choice = []
count_num = 0
none_list = []
none_action = False

while action:
    a_choice = input('请输入一个想要对比的序号，不得超过{},不得小于{},输入q或者Q表示退出选择:'.format(len(collector)-1,0)) 
    if a_choice == 'q' or a_choice == 'Q':

        #非空检验
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



#数据对应字典

name_dict = {'1':'内外向(E)', '2':'神经质(N)','3':'精神质(P)', '4':'掩饰性(L)',\
        '5':'躯体化', '6':'强迫症状', '7':'人际关系敏感', '8':'抑郁', '9':'焦虑', '10':'敌对', '11':'恐怖', '12':'偏执', '13':'精神病性', '14':'其他', '15':'总分', '16':'总均分', '17':'阳性项目数'}
list_match = []

#选择对比数据
action = True

while  action:
    a_match = input("\n1.内外向(E),\n 2.神经质(N),\n3.精神质(P),\n 4.掩饰性(L),\n5.躯体化,\n 6.强迫症状,\n 7.人际关系敏感,\n 8.抑郁,\n 9.焦虑,\n 10.敌对,\n 11.恐怖,\n 12.偏执,\n 13.精神病性,\n 14.其他,\n 15.总分,\n 16.总均分,\n 17.阳性项目数\n请选择一个要对比的指标（输入q或者Q表示推出选择），输入序号就行：")
    if a_match == 'q' or a_match == 'Q':
        break
    else:
        list_match.append(a_match)


#x点初始数据


#获取数据
x = []
y_data = 1
for num in list_choice:
    m = 1
    n = 2
    name = collector[0]
    guy = collector[num] 
    x_name = []
    z0 = []

    for n_ma in list_match:
        x_name.append(name_dict[n_ma])
        z0.append(guy[name.index(name_dict[n_ma])])

   

#3d柱体数据设定
    y1 = y3 = int(y_data)-0.5
    y2 = y4 = int(y_data)
    for z_data in z0:
        color3D = randomcolor()

        x.append((m+n)/2)
        h = float(z_data)
        x1 = x2 = m
        x3 = x4 = n
        m += 1.5
        n += 1.5
        bar_of_3d(x1,x2,x3,x4,y1,y2,y3,y4,h,color3D)
    y_data += 1
        

#规范坐标轴
ax1.set_xlabel('属性名称',labelpad=22)
ax1.set_ylabel('对比序号及性别',labelpad=18)
ax1.set_zlabel('所选值')

#替换y轴
y_lname = list(map(int,list_choice))
y_nname = []
for n in y_lname:
    sex_index = name.index('性别')
    sex = collector[n][sex_index]
    y_nname.append('序号'+str(n)+'性别'+sex)

y_lname = [x+1 for x in range(len(list_choice))]
yticks(y_lname,y_nname,rotation=-10)

#替换x轴

xticks(x,x_name,fontsize=8,rotation=50)

title('自选人数自选属性对比3d直方图')


show()
fig.savefig('自选人数自选属性对比3d直方图')