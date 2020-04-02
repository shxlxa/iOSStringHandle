#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xlwt
import xlutils.copy



# 将excel表sheet1中的中文数据与sheet0中的中文进行对比，中文相同时
# 将sheet1中的key替换成sheet1的key,写入到一个新的excel表中

# 1.写入android的key到新的excel文档
# 2.写入iOS所有数据到新的excel文档
def handle_iOS():
	global data3
	global sheet3
	data = xlrd.open_workbook('origin.xls') # 打开xls文件
	sheet0 = data.sheets()[0] # 打开第一张表

	sheet1 = data.sheets()[1] # 打开第二张表
	nrows = sheet1.nrows      # 获取表的行数
	ncols = sheet1.ncols		 # 获取表的列数

	# 创建一个新的excel文档
	data3 = xlwt.Workbook() 
	sheet3 = data3.add_sheet('sheet3',cell_overwrite_ok=True)
	for row in range(sheet0.nrows):
		# 依次读取sheet1第一列的数据（en数据）
		cell = sheet0.cell_value(row,2)
		print(cell)
		# 依次读取sheet1第一列的数据	
		for rowB in range(sheet1.nrows):
			cellB = sheet1.cell_value(rowB,2)
			if cell == cellB:
				print 'row '+str(rowB)+' '+cell
				# print 'row ' + str(row) + ' ' + sheet1.cell_value(rowB,0)
				# 1.将android_key写入到第一列中
				android_key = sheet1.cell_value(rowB,0)
				sheet3.write(row,0,android_key)
				data3.save('osd.xls')
		# 2.依次将sheet1中的数据写入到新的excel文件中（从第二列开始）
		for col in range(0,sheet0.ncols):
			print 'row '+str(row)+' col '+str(col)
			cell = sheet0.cell_value(row,col)
			sheet3.write(row,col+1,cell)
			data3.save('osd.xls')


# 处理安卓和iOS中文对比不同的部分，加到下面
def handle_android():
	data = xlrd.open_workbook('origin.xls') # 打开xls文件
	sheet0 = data.sheets()[0] # 打开第一张表
	nrows = sheet0.nrows      # 获取表的行数
	ncols = sheet0.ncols		 # 获取表的列数

	sheet1 = data.sheets()[1] # 打开第二张表

	global data3
	global sheet3
	rowNum = sheet0.nrows
	# 以行来遍历sheet1中的数据，依次取出第2列的数据cell
	for row in range(sheet1.nrows):

		flag = 0
		cell = sheet1.cell_value(row,2)
		# print cell
		# 以行来遍历sheet0中的数据，依次取出第2列的数据cellB
		# 比较cell和cellB,相等的话flag置1
		for rowB in range(sheet0.nrows):
			cellB = sheet0.cell_value(rowB,2)
			if cell == cellB:
				flag = 1
		# 依次取出sheet1中第一列的数据key
		# 如果flag=0并且key不为空，则将数据依次添加到sheet4中，最后写到android.xls中
		key = sheet1.cell_value(row,0)
		if flag==0 and key != '':
			rowNum += 1
			# print sheet1.cell_value(row,0) + 'row '+str(rowNum)
			# 中间的第1列是留给iOS_key的
			sheet3.write(rowNum,0,sheet1.cell_value(row,0))
			sheet3.write(rowNum,2,sheet1.cell_value(row,1))
			sheet3.write(rowNum,3,sheet1.cell_value(row,2))

			sheet3.write(rowNum,4,sheet1.cell_value(row,3))
			sheet3.write(rowNum,5,sheet1.cell_value(row,4))
			sheet3.write(rowNum,6,sheet1.cell_value(row,5))
			sheet3.write(rowNum,7,sheet1.cell_value(row,6))
			sheet3.write(rowNum,8,sheet1.cell_value(row,7))
			sheet3.write(rowNum,9,sheet1.cell_value(row,8))
			sheet3.write(rowNum,10,sheet1.cell_value(row,9))
			sheet3.write(rowNum,11,sheet1.cell_value(row,10))
			sheet3.write(rowNum,12,sheet1.cell_value(row,11))
			sheet3.write(rowNum,13,sheet1.cell_value(row,12))

			sheet3.write(rowNum,14,sheet1.cell_value(row,13))
			sheet3.write(rowNum,15,sheet1.cell_value(row,14))
			sheet3.write(rowNum,16,sheet1.cell_value(row,15))
			data3.save('osd.xls')
	
			
handle_iOS()
handle_android()




