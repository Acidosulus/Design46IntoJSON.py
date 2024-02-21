import sys
from openpyxl import load_workbook
from click import echo, style

import pprint
printer = pprint.PrettyPrinter(indent=12, width=180)
prnt = printer.pprint

echo(style(text=sys.argv[1], bg='red', fg='bright_yellow'))

def join_tuple_into_string(source:tuple):
	result = ''
	for elem in source:
		if type(elem)==str:
			result += elem
		if type(elem)==tuple:
			result += join_tuple_into_string(elem)
	return result

def read_excel_range(file_name, sheet_name, start_row, end_row, start_col, end_col):
	import win32com.client as win32
	excel = win32.Dispatch('Excel.Application')
	workbook = excel.Workbooks.Open(file_name)
	sheet = workbook.Sheets[sheet_name]
	range_cells = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
	data = []
	for row in range_cells.Rows:
		row_data = []
		for cell in row:
			if not cell.Hidden:
				if cell.Value2!=None:
					if type(cell.Value2)==tuple:
						if  len(cell.Value2)>0:
							val = cell.Value2
							val = join_tuple_into_string(val)
							row_data.append(val)
					else:
						row_data.append(str(cell.Value2))
		if len(row_data)>0:
			data.append(''.join(row_data))
	workbook.Close(False)
	excel.Quit()
	return '\n'.join(data)



pattern_content = open(file='pattern.json', mode='r', encoding='utf-8').read()



echo(style('Титульный', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Титульный }}',read_excel_range(sys.argv[1],'Титульный', 5, 71, 11, 11))

echo(style('Раздел_I__А', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_I__А }}',read_excel_range(sys.argv[1],'Раздел_I__А', 18, 94, 105, 198))

echo(style('Раздел_I__Б', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_I__Б }}',read_excel_range(sys.argv[1],'Раздел_I__Б', 18, 94, 19, 19))

echo(style('Раздел_I__В', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_I__В }}',read_excel_range(sys.argv[1],'Раздел_I__В', 18, 94, 28, 44))

echo(style('Раздел_II__А_(ТИС)', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_II__А_(ТИС) }}',read_excel_range(sys.argv[1],'Раздел_II__А_(ТИС)', 18, 39, 67, 123))

echo(style('Раздел_II__Б_(ТИС)', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_II__Б_(ТИС) }}',read_excel_range(sys.argv[1],'Раздел_II__Б_(ТИС)', 18, 39, 15, 15))

echo(style('Раздел_III', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_III }}',read_excel_range(sys.argv[1],'Раздел_III', 18, 45, 12, 12))

echo(style('Раздел_IV', bg='yellow', fg='bright_blue'))
pattern_content = pattern_content.replace('{{ Раздел_IV }}',read_excel_range(sys.argv[1],'Раздел_IV', 18, 41, 12, 12))


open(file=sys.argv[1]+'.json', mode='w', encoding='utf-8').write(pattern_content)




