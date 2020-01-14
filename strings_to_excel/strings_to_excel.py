#!/usr/bin/python
import xlwt as ExcelWrite
import codecs
import sys
import xlrd
import xlwt
import time

# python2.7 strings_to_excel.py

reload(sys)
sys.setdefaultencoding('utf-8')


def write_file(xls, sheet, key_list, string_name, index):
	file = codecs.open(string_name, 'r', 'utf-8')
	lines = file.readlines()
	# print(lines)
	file.close()
	sheet.write(0,index,string_name.split('.')[0])
	item_num = 0
	for line in lines:
	    new_line = line.strip('\n')
	    if len(new_line) <= 0:
	        continue
	    
	    if new_line.startswith('\"'):
	        item_num += 1
	        ios_key = new_line.split('\"')[1]
	        ios_value = new_line.split('=')[1].split('\"')[1]
	        if index==1:
	        	sheet.write(item_num, 0, ios_key)
	        	sheet.write(item_num, index, ios_value)
	        else:
	        	# Traverse the original ios key_list, compare the current key a_value with it, 
	        	# and write the value to the row corresponding to the key if the key is the same
	        	for (a_index, a_value) in enumerate(key_list): #
	        		if a_value == ios_key:
	        			sheet.write(a_index, index, ios_value)
	        			print ios_value
	        			break
	xls.save("1.xls")


def write_multi_strings():
	xls = ExcelWrite.Workbook()
	sheet = xls.add_sheet("sheet1")
	write_file(xls, sheet, [], "cn.strings", 1)
	
	key_list = readXmls()

	languagesStrings = ["en.strings","tr.strings","fr.strings"]
	index = 1
	for mString in languagesStrings:
		index += 1
		write_file(xls, sheet, key_list, mString, index)

def readXmls():
    data = xlrd.open_workbook('1.xls')
    key_list = []
    sheet1 = data.sheets()[0]
    nrows = sheet1.nrows
    for row in range(nrows):
        cellValue = sheet1.cell_value(row,0)
        key_list.append(cellValue)
    return key_list

write_multi_strings()

	