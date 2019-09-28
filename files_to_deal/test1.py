string1 = "星期一星期二星期三星期四星期五星期六星期日"
a = eval(input("请输入阿拉伯数字对应星期："))
print(string1[(a - 1)*3:(a - 1)*3+3])
