import openpyxl
from openpyxl import Workbook, load_workbook
#from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import re
import csv


wb = openpyxl.Workbook()
ws = wb.active

#Конвертируем csv в xlsx
with open('main.csv') as f:
    reader = csv.reader(f, delimiter=';')
    for row in reader:
        ws.append(row)
wb_name='main.xlsx'

wb.save(wb_name)
#
wb = load_workbook(wb_name)
ws = wb.active
col_len=len(ws['A:A'])
color_good="0099CC00"
color_bad="00993300"
color_not_good_not_bad="00FF9900"

#	D:D 	FIX STATUS GREEN-RED
col_range = ws['D:D']
for i in range(1,col_len):
	if col_range[i].value == 1: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)

#	E:E 	FLOAT STATUS GREEN-RED
col_range = ws['E:E']
for i in range(1,col_len):
	if col_range[i].value == 0: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)

#	F:F 	GPS STATUS GREEN-RED
col_range = ws['F:F']
for i in range(1,col_len):
	if col_range[i].value == 0: col_range[i].font = Font(color=color_good, bold="True")
	else : col_range[i].font = Font(color=color_bad)

#	G:G 	DGPS STATUS GREEN-RED
col_range = ws['G:G']
for i in range(1,col_len):
	if col_range[i].value == 0: col_range[i].font = Font(color=color_good, bold="True")
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

	if val < 7.1: col_range[i].font = Font(color=color_bad)
	elif val >=7.1 and val < 7.4: col_range[i].font = Font(color=color_not_good_not_bad)
	else: col_range[i].font = Font(color=color_good)

	#for i in range(1,col_len):
#m = re.search(r'\d{1,}\.\d{1,} ','157.932 [10]')	
#m = re.search(r'\S{1,}\s', '157.932 [10]')
#print(float(m[0]))
#print(m[0] if m else 'Not found')
		# 157.9 [XY:127.8 Z: 92.7]
	
	# Пока обрезаю для теста вручную
wb.save(wb_name)
