from json import load
import sys
import os
import openpyxl

def GetSheetCount(path:str) -> int:
	"""Get a number of sheet.

	Get a number of sheets in a file specified by argument path.

	Args:
		path(str):	Path to Excel file whose extention is xlsx to check the number of sheet.

	Retunrs:
		The number of sheets the excel file has.
	
	"""
	wb = openpyxl.load_workbook(path)
	sheet_num = len(wb.worksheets)
	return sheet_num

if '__main__' == __name__:
	path = sys.argv[1]
	sheet_num = GetSheetCount(path)

	print('sheet num -', sheet_num)
