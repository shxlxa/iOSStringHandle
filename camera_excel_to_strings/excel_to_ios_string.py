
#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlwt
import codecs
import xlrd
import sys

# python excel_to_ios_string.py


file = xlrd.open_workbook('osd.xls')
table = file.sheets()[0]

nrow = table.nrows
ncol = table.ncols

def saveFile(index_num,fileName):
    f = open(fileName,'w')

    for i in range(0,nrow):
        data_key = table.cell(i,0).value
        # iOS key为空的话，不进行写操作
        if data_key != '':
	        data_value = table.cell(i,index_num).value
	        data_a = '\"' + str(data_key) + '\"' + ' = ' + '\"' + str(data_value) + '\";'
	       
	        f.write(data_a)
	        f.write('\n')
	        print(data_a)
    f.close()


def saveStrings():
	name_arr = ['en.strings', 'zh.strings', 'zh_tr.strings', 'baojia.strings', 'et.strings', 'lt.strings', 'lv.strings', 'ro.strings', 'uk.strings', 'be.strings', 'ru.strings']
	for value in name_arr:
		idx = name_arr.index(value)
		saveFile(idx+2, value)

saveStrings()


