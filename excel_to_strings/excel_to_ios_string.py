
#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlwt
import codecs
import xlrd
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

file = xlrd.open_workbook('osd.xls')
table = file.sheets()[0]

nrow = table.nrows
ncol = table.ncols

def saveFile(index_num,fileName):
    f = open(fileName,'w')

    for i in range(0,nrow):
        for j in range(0,3):
            data_en = table.cell(i,0).value
            data_ch = table.cell(i,index_num).value
        
        data_a = '\"' + str(data_en) + '\"' + ' = ' + '\"' + str(data_ch) + '\";'
       
        f.write(data_a.encode('utf-8'))
        f.write('\n')
        print(data_a.encode('utf-8'))
    f.close()


def saveStrings():
	name_arr = ['en.strings', 'Italian.strings', 'spanish.strings', 'French.strings', 'German.strings', 'Dutch.strings']
	for value in name_arr:
		idx = name_arr.index(value)
		saveFile(idx+1, value)

saveStrings()


