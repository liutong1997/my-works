# coding = utf-8
from multiprocessing import Process
import docx


# 定义替换名字并记录编号的函数
def replace_name_and_count(t):
    # 引用全局变量
    global a
    global c
    # 用字典储存编号
    name[t.cell(0, 3).text + t.cell(0, 1).text] = a
    # 编号自增
    a += 1
    # 将编号转换为字符串
    b = str(name.get(t.cell(0, 3).text + t.cell(0, 1).text))
    # 清除名字
    t.cell(0, 3).paragraphs[0].clear()
    # 用编号替换名字
    t.cell(0, 3).paragraphs[0].add_run(b)
    # 清楚原‘姓名’字符串，替换为‘编号’
    t.cell(0, 2).paragraphs[0].clear()
    t.cell(0, 2).paragraphs[0].add_run(c)


def replace_name(t):
    # 引用全局变量
    global c
    # 将编号转换为字符串
    b = str(name.get(t.cell(0, 3).text + t.cell(0, 1).text))
    # 清除名字
    t.cell(0, 3).paragraphs[0].clear()
    # 用编号替换名字
    t.cell(0, 3).paragraphs[0].add_run(b)
    c = '序号'
    # 清楚原‘姓名’字符串，替换为‘编号’
    t.cell(0, 2).paragraphs[0].clear()
    t.cell(0, 2).paragraphs[0].add_run(c)


if __name__ == '__main__':
    # 输入正确的文件路径
    file = docx.Document(input('请输入要读取的正确的docx文件路径:'))
    # 读取所有表格
    tables = file.tables
    # 定义全局变量c,a分别存放‘序号’字符串和
    c = '序号'
    a = 1
    # 定义字典储存姓名编号
    name = {}
    # 遍历所有表，迭代得出每一张表
    for i in range(len(tables)):
        # 将每一张表传递给参数t
        t = tables[i]
        # 查找目标表
        if t.cell(0, 0).text == '编号':
            # 判断是否已经收录了个人信息
            if t.cell(0, 3).text + t.cell(0, 1).text not in name:
                # Process 对象只是一个子任务，运行该任务时系统会自动创建一个子进程
                p = Process(target=replace_name_and_count(t))
                # 启动一个子进程来运行子任务
                p.start()
                p.join()
                # 子进程完成后，继续运行主进程,保存文件
            else:
                p = Process(target=replace_name(t))
                # 启动一个子进程来运行子任务
                p.start()
                p.join()
                # 子进程完成后，继续运行主进程,保存文件
    file.save(input('请输入要写入的正确的docx文件路径，如果与读取路径重名将覆盖源文件，请慎重:'))
