# import pandas as pd
# import numpy as np
#
# # s = pd.Series([1, 3, 6, np.nan, 44, 1])
# # dates = pd.date_range('20160101', periods=6)
# # df = pd.DataFrame(np.random.randn(6, 4), index=dates, columns=['a', 'b', 'c', 'd'])
#
# df2 = pd.DataFrame({'A': 1.,
#                     'B': pd.Timestamp('20130102'),
#                     'C': pd.Series(1, index=[range(4)], dtype='float32').values,
#                     'D': np.array([3] * 4, dtype='int32'),
#                     'E': pd.Categorical(["test", "train", "test", "train"]),
#                     'F': 'foo'}, index=list('abcd'))
#
# # print(pd.Series(1, index=list(range(4)), dtype='float32'))
# print(df2)


# coding=utf-8
import os
from multiprocessing import Process


def hello(name):
    # os.getpid() 用来获取当前进程 ID
    print('child process: {}'.format(os.getpid()))
    print('Hello ' + name)


def main():
    # 打印当前进程即主进程 ID
    print('parent process: {}'.format(os.getpid()))
    # Process 对象只是一个子任务，运行该任务时系统会自动创建一个子进程
    # 注意 args 参数要以 tuple 方式传入
    p = Process(target=hello, args=('shallot',))
    print('child process start')
    # 启动一个子进程来运行子任务，该进程运行的是 hello() 函数中的代码
    p.start()
    p.join()
    # 子进程完成后，继续运行主进程
    print('child process stop')
    print('parent process: {}'.format(os.getpid()))


if __name__ == '__main__':
    main()
