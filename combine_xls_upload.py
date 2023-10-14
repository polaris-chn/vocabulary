import xlwt
import xlrd
from xlutils.copy import copy
import os
target_path = 'word_combine.xls'
source_path1 = 'CET4_edited.xls'
source_path2 = 'CET6_edited.xls'
source_path3 = 'TOEFL.xls'

if os.path.exists(target_path):
	os.remove(target_path)
	book = xlwt.Workbook()
	sheet = book.add_sheet('sheet1')
	book.save(target_path)
else:
	book = xlwt.Workbook()
	sheet = book.add_sheet('sheet1')
	book.save(target_path)


def word_write(source_book, target_book):
	book1 = xlrd.open_workbook(source_book)
	sheet1 = book1.sheets()[0]
	rows1 = sheet1.nrows
	cols1 = sheet1.ncols

	book2 = xlrd.open_workbook(target_book)
	sheet2 = book2.sheets()[0]
	rows2 = book2.sheets()[0].nrows
	cols2 = book2.sheets()[0].ncols

	book3 = copy(book2)
	sheet3 = book3.get_sheet(0)

	# write in when target book is empty
	if rows2 == 0:
		for i in range(rows1):
			for j in range(cols1):
				sheet3.write(rows2+i, j, sheet1.cell_value(i, j))
	# write in when target book is not empty
	else:
		source_list = []
		target_list = []
		row_index = []
		for i in range(rows1):
			source_list.append(sheet1.cell_value(i, 0))
		for k in range(rows2):
			target_list.append(sheet2.cell_value(k, 0))
		# find item and index that in source list and not in target list
		retD = list(set(source_list).difference(set(target_list)))
		for value in retD:
			row_index.append(source_list.index(value))
			
		for j in range(cols1):
			for a in range(len(row_index)):
				sheet3.write(rows2+a, j, sheet1.cell_value(row_index[a], j))

	book3.save(target_book)

word_write(source_path1, target_path)
word_write(source_path2, target_path)
word_write(source_path3, target_path)