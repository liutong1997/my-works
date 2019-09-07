# coding = utf-8
import docx
from multiprocessing import Process


def deal_alltable(t):
    # 判断是否为含有编号的单元格，含有即为用户信息表
    if t.cell(0, 0).text == '编号':
        # 创建变量m收集行数，每张表进行一次初始化
        m = 0

        # 遍历行
        for i in range(len(t.rows)):
            # 创建变量n收集列数，每行进行一次初始化
            n = 0
            # 统计行数
            m += 1
            # 每遍历一次行遍历该行中所有列
            for j in range(len(t.rows[i].cells)):
                # 统计列数
                n += 1
                # 判断如果为第二行则开始清洗内容
                if i >= 1:
                    # 清空段落数据
                    run = t.cell(i, j).paragraphs[0].clear()
        # 合并单元格
        t.cell(1, 0).merge(t.cell(m - 1, n - 1))


if __name__ == '__main__':
    # 输入正确的文件路径
    file = docx.Document(input('请输入要读取的正确的docx文件路径:'))
    # 读取所有表格
    tables = file.tables
    # 遍历所有表，迭代得出每一张表
    for i in range(len(tables)):
        # 将每一张表传递给参数t
        t = tables[i]
        # Process 对象只是一个子任务，运行该任务时系统会自动创建一个子进程
        p = Process(target=deal_alltable(t))
        # 启动一个子进程来运行子任务
        p.start()
        p.join()
        # 子进程完成后，继续运行主进程,保存文件
    file.save(input('请输入要写入的正确的docx文件路径，如果与读取路径重名将覆盖源文件，请慎重:'))
