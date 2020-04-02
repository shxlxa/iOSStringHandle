#!/usr/bin/python
import xlwt as ExcelWrite
import codecs
import sys
import xlrd
import xlwt
import time

# python strings_to_excel.py

def write_file(xls, sheet, key_list, string_name, col_num):
	file = codecs.open(string_name, 'r', 'utf-8')
	lines = file.readlines()
	# print(lines)
	file.close()
	sheet.write(0,col_num,string_name.split('.')[0])
	row_num = 0
	for line in lines:
	    new_line = line.strip('\n')
	    if len(new_line) <= 0:
	        continue
	    
	    if new_line.startswith('\"'):
	        row_num += 1
	        ios_key = new_line.split('\"')[1]
	        ios_value = new_line.split('=')[1].split('\"')[1]
	        # 写第1列的数据（中文）,作为后面写入的基准
	        if col_num==1:
	        	sheet.write(row_num, 0, ios_key)
	        	sheet.write(row_num, col_num, ios_value)
	        else:
	        	# 遍历key_list,当ios_key与遍历的key相同时，将值写到这一行
	        	for (a_index, a_value) in enumerate(key_list): 
	        		if a_value == ios_key:
	        			sheet.write(a_index, col_num, ios_value)
	        			# print(ios_value)
	        			break
	xls.save("1.xls")


def write_multi_strings():
	# 1.将cn.string中的key和value写到excel文件1.xls中
	xls = ExcelWrite.Workbook()
	sheet = xls.add_sheet("sheet1")
	write_file(xls, sheet, [], "cn.strings", 1)
	
	# 2.读取1.xls中key的列表，保存在数组key_list中
	key_list = readXmls()
	print(key_list)

	# 3.写入其他语言的string文件，写入时会将该string语言的key与key_list对比，保证写到1.xls
	#   中相同key的那一行里
	languagesStrings = ["en.strings","tr.strings","fr.strings"]
	col_num = 1
	for mString in languagesStrings:
		col_num += 1
		write_file(xls, sheet, key_list, mString, col_num)

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

	