import sys
import openpyxl
from openpyxl import worksheet
import SetFooter as sf

def SetFooter(path:str, footers:list) -> None:
	"""Set footer of left.

	Setup left footer of all sheets in a excel file.

	Args:
		path(str): Path to file to set footer.
		headetrs(list): Collection of strings to set into footer.
						One item is one line in footer.
						When output, all items are joined by change line code.
	"""
	wb = openpyxl.load_workbook(path)

	sheets = wb.worksheets
	for sheet in sheets:
		SetFooterToSheet(sheet=sheet, footers=footers)

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
	footer_part = sheet.oddFooter.left
	sf.SetFooter(footer_part=footer_part, footers=footers)

if '__main__' == __name__:
	path = sys.argv[1]
	argc = len(sys.argv)
	footers = sys.argv[2:argc]
	SetFooter(path=path, footers=footers)
