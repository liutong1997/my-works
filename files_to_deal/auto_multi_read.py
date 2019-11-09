import os
import glob
import shutil


def make_folder(p):
    if os.path.exists(p):
        shutil.rmtree(p)
    os.mkdir(p)


file_path = os.getcwd()
local_path = file_path + '/omg'
make_folder(local_path)
for file in glob.glob(os.path.join(file_path, '*.docx')):
    print(file_path)
    print(file)
    print(local_path)
    if local_path[0] == '/':
        print(local_path+'/'+file.strip(file_path))
    else:
        print(file_path + '\omg' + '\\' + file.strip(file_path))


