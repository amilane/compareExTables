import openpyxl
import os
from operator import itemgetter
from itertools import groupby

os.chdir('D:\\CompareTables')
# ИМЯ ФАЙЛА
filename = 'comparison  09 2018'

wb = openpyxl.load_workbook(filename + '.xlsx')

# ЛИСТЫ
overSheet = wb.get_sheet_by_name('overall list of details')
lisapSheet = wb.get_sheet_by_name('LISAP')
resultSheet = wb.create_sheet('Result')
notInOverSheet = wb.create_sheet('Not in Overall')

# ДИАПАЗОНЫ В ЛИСТАХ
overRange = range(4, 11285)
lisapRange = range(2, 70081)

# сбор данных для сравнения
overID = [overSheet.cell(row = i, column = 15).value for i in overRange]
lisID_NUM = [[lisapSheet.cell(row = i, column = 3).value, lisapSheet.cell(row = i, column = 9).value] for i in lisapRange if lisapSheet.cell(row = i, column = 3).value != "" and lisapSheet.cell(row = i, column = 9).value != ""]

# уникальные значения
uniqListID_NUM = []
for i in lisID_NUM:
  if i not in uniqListID_NUM:
    uniqListID_NUM.append(i)

# сортировка-группировка по ID
sortedUniqListID_NUM = sorted(uniqListID_NUM, key = lambda i: i[0])
groupID_NUM = [[x for x in g] for k,g in groupby(sortedUniqListID_NUM, lambda i: i[0])]

# схлопывание номеров накладных в 1 строку
flatGroupID_NUM = []
for group in groupID_NUM:
	a = []
	a.append(group[0][0])
	b = ""
	for g in group:
		b += g[1] + ", "
	a.append(b)
	flatGroupID_NUM.append(a)

# список накладных в порядке ID из overall
lisID_in = ['' for i in overID]
lisID_out = []
for i in flatGroupID_NUM:
	if i[0] in overID:
		ind = overID.index(i[0])
		lisID_in[ind] = i[1]
	else:
		lisID_out.append(i)

# вывод данных и сохранение нового документа
for i, k in zip(overRange, lisID_in):
  resultSheet.cell(row = i, column = 21).value = k

for i in range(1, len(lisID_out)+1):
  notInOverSheet.cell(row = i, column = 1).value = lisID_out[i-1][0]
  notInOverSheet.cell(row = i, column = 2).value = lisID_out[i-1][1]


wb.save(filename + '_2.xlsx')

print("OK")