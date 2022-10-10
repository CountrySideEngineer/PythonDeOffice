import sys
import openpyxl
from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart
import SetFooter as sf

def SetFooter(path:str, footer:list) -> None:
	"""Set header of right.

	Setup right header of all sheets in a excel file.

	Args:
		path(str): Path to file to set header.
		headetrs(list): Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	wb = openpyxl.load_workbook(path)

	sheets = wb.worksheets
	for sheet in sheets:
		SetFooterToSheet(sheet=sheet, footers=footer)

	wb.save(path)
	wb.close()

def SetFooterToSheet(sheet:worksheet, footers:list) -> None:
	"""Set footer of center.

	Setup center footer of all sheets in a excel file.

	Args:
		sheet(worksheet): Worksheet object to set header.
		headetrs(list): Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	footer_part = sheet.oddFooter.right
	sf.SetFooter(footer_part=footer_part, footers=footers)

if '__main__' == __name__:
	path = sys.argv[1]
	argc = len(sys.argv)
	footers = sys.argv[2:argc]
	SetFooter(path=path, footer=footers)
