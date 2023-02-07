import openpyxl
from openpyxl import Workbook, load_workbook
#from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import re #регулярки для поиска кусков строки в ячейке
import csv #будем конвертировать csv в xlsx
import os #чисто проверим нужные библиотеки перед запуском
requirementsFile = 'requirements.txt'
"""
if os.path.isfile(requirementsFile):
	os.system('pip3 install -r %s' % requirementsFile)
else:
	print('File "%s" not found' % requirementsFile)
"""

wb = openpyxl.Workbook()
ws = wb.active

#Конвертируем csv в xlsx
with open('main.csv') as f:
    reader = csv.reader(f, delimiter=';')
    for row in reader:
        ws.append(row)
wb_name='main.xlsx'
wb.save(wb_name) # По крайней мере файлик надо сохранить
				 # а как уже дальше пойдет- посмотрим
wb = load_workbook(wb_name)
ws = wb.active
col_len=len(ws['A:A'])
color_good="0099CC00"
#color_bad="00993300"
color_bad="FF0000"
color_not_good_not_bad="00FF9900"

#	D:D 	FIX STATUS GREEN-RED
col_range = ws['D:D']
for i in range(1,col_len):
	val=float(re.search(r'\d{1,}',col_range[i].value)[0])
	if val == 100: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)
	#print(val[0])

#	E:E 	FLOAT STATUS GREEN-RED
col_range = ws['E:E']
for i in range(1,col_len):
	val=float(re.search(r'\d{1,}',col_range[i].value)[0])
	#print('i=', i,'val=',val)

	if val == 0: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)

#	F:F 	GPS STATUS GREEN-RED
col_range = ws['F:F']
for i in range(1,col_len):
	val=float(re.search(r'\d{1,}',col_range[i].value)[0])
	if val == 0: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)

#	G:G 	DGPS STATUS GREEN-RED
col_range = ws['G:G']
for i in range(1,col_len):
	val=float(re.search(r'\d{1,}',col_range[i].value)[0])
	if val == 0: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)


#	H:H 	RTCM
col_range = ws['H:H']
for i in range(1,col_len):
	#val=float(col_range[i].value)
	val = float(re.search(r'\S{1,}\s',col_range[i].value)[0])
	#print(val)
	#print("i=",i,":",val)
	if val < 1.9: col_range[i].font = Font(color=color_good)
	elif val >=1.9 and val < 4: col_range[i].font = Font(color=color_not_good_not_bad)
	else: col_range[i].font = Font(color=color_bad)

#	H:H 	VOLTAGE
col_range = ws['K:K']
for i in range(1,col_len):
	volt=col_range[i].value
	val=float(volt.split('-')[1][0:-1])
	#print("i=",i,":",val)

	if val < 7.16: col_range[i].font = Font(color=color_bad)
	elif val >=7.16 and val < 7.5: col_range[i].font = Font(color=color_not_good_not_bad)
	else: col_range[i].font = Font(color=color_good)

#	L:L 	RSSI
col_range = ws['L:L']
for i in range(1,col_len):
	val=float(col_range[i].value.split(' ')[1])
	#val = float(re.search(r'\S{1,}\s',col_range[i].value)[0])
	#print(val)
	#print("i=",i,":",val)
	if val >= -70: col_range[i].font = Font(color=color_good)
	elif val >=-80 and val < -70: col_range[i].font = Font(color=color_not_good_not_bad)
	else: col_range[i].font = Font(color=color_bad)

def check_accuracy(array):
	for i in range(1,col_len):
		#result = array[i].value.split(' ')
		result = re.search(r'\d{1,}',array[i].value)
		print(result[0])
		"""
		val_average = float(result[2])
		val_max = float(re.search(r'\f{1,}',result[4])[0])
		if val_average < 20 and val_max <100: array[i].font = Font(color=color_good)
		#elif val >=-80 and val < -70: col_range[i].font = Font(color=color_not_good_not_bad)
		else: array[i].font = Font(color=color_bad)
		"""

#	for both h_accuracy and v_accuracy
#	requirements are the same and listed 
#	in check_accuracy function

#	Q:Q 	V_ACCUR
check_accuracy(ws['Q:Q'])
#	R:R 	H_ACCUR
check_accuracy(ws['R:R'])

"""for i in range(1,col_len):
	result = col_range[i].value.split(' ')
	val_average = float(result[2])
	val_max = float(re.search(r'\d{1,}',result[5])[0])
	#val = float(re.search(r'\S{1,}\s',col_range[i].value)[0])
	#print(val_max)
	#print("i=",i,":",val)
	if val_average < 20 and val_max <100: col_range[i].font = Font(color=color_good)
	#elif val >=-80 and val < -70: col_range[i].font = Font(color=color_not_good_not_bad)
	else: col_range[i].font = Font(color=color_bad)
"""


	# ----TESTS-----
	#for i in range(1,col_len):
#m = re.search(r'\d{1,}\.\d{1,} ','157.932 [10]')	
#m = re.search(r'\S{1,}\s', '157.932 [10]')
#print(float(m[0]))
#print(m[0] if m else 'Not found')
		# 157.9 [XY:127.8 Z: 92.7]
	
	# Пока обрезаю для теста вручную
#print(os.time())
wb.save(wb_name)
