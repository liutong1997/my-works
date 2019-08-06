import csv
import os
import glob
from datetime import date
from xlrd import open_workbook,xldate_as_tuple
item_numbers_file = 'items.csv'
path_to_folder = r'D:\pythonworks\files_to_deal\files'
output_file = 'output.csv'
item_numbers_to_find=[]
with open(item_numbers_file,'r',newline='')as item_numbers_csv_file:
    filereader = csv.reader(item_numbers_csv_file)
    for row in filereader:
        item_numbers_to_find.append(row[0])
print(item_numbers_to_find)
filewriter = csv.writer(open(output_file,'a',newline=''))
file_count = 0
line_count = 0
item_numbers_count = 0
for input_file in glob.glob(os.path.join(path_to_folder,'*.*')):
    file_count+=1
    if input_file.split('.')[-1] == 'csv':
        with open(input_file,'r',newline='') as csv_in_file:
            filereader = csv.reader(csv_in_file)
            header = next(filereader)
            for row in filereader:
                row_of_output = []
                for column in range(len(header)):
                    if column == 3:
                        cell_value = str(row[column]).lstrip('$').replace(',','').strip()
                        row_of_output.append(cell_value)
                    else:
                        cell_value = str(row[column]).strip()
                        row_of_output.append(cell_value)
                row_of_output.append(os.path.basename(input_file))
                if row[0] in item_numbers_to_find:
                    filewriter.writerow(row_of_output)
                    item_numbers_count+=1
                line_count+=1
    elif input_file.split('.')[-1] == 'xls' or input_file.split('.')[-1] == 'xlsx':
        workbook = open_workbook(input_file)
        for worksheet in workbook.sheets():
            try:
                header = worksheet.row_values(0)
            except IndexError:
                pass
            for row in range(1,worksheet.nrows):
                row_of_output = []
                for column in range(len(header)):
                    if worksheet.cell_type(row,column) == 3:
                        cell_value = xldate_as_tuple(worksheet.cell_value(row,column),workbook.datemode)
                        cell_value = str(date(*cell_value[0:3])).strip()
                        row_of_output.append(cell_value)
                    else:
                        cell_value = str(worksheet.cell_value(row,column)).strip()
                        row_of_output.append(cell_value)
                row_of_output.append(os.path.basename(input_file))
                row_of_output.append(worksheet.name)
                if str(worksheet.cell_value(row,0)).split('.')[0].strip() in item_numbers_to_find:
                    filewriter.writerow(row_of_output)
                    item_numbers_count += 1
                line_count += 1
print('Number of files:',file_count)
print('number of lines:',line_count)
print('number of items:',item_numbers_count)