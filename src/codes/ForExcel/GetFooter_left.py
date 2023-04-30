import sys
import openpyxl
from openpyxl import worksheet
import GetFooter as gf

def GetFooter(path:str) -> list:
	"""Get footer of left side.
	
	"""
	wb = openpyxl.load_workbook(path)

	footer_list = []
	sheets = wb.worksheets
	for sheet in sheets:
		footer_item = GetFooterFromSheet(sheet)
		footer_list.append(footer_item)

	wb.close()

def  GetFooterFromSheet(sheet:worksheet) -> str:
	footer_part = worksheet.oddFooter.left
	footer = gf.GetFooter(footer_part=footer_part)

	return footer
