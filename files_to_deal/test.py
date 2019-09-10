import os
import glob
import shutil


def make_folder(p):
    if os.path.exists(p):
        shutil.rmtree(p)
    os.mkdir(p)


file_path = os.getcwd()
for file in glob.glob(os.path.join(file_path, 'new*.docx')):
    print(file_path)
    print(file)
    print(file.strip(file_path))

local_path = file_path + '/omg'
make_folder(local_path)
