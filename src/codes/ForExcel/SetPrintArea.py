from ctypes.wintypes import HPALETTE
import sys
import openpyxl
from openpyxl import worksheet

def SetPrintAreaAuto(path:str) -> None:
	"""Set the print area into AUTO.
	
	Setup the print area of all sheets in a file specified by argument into AUTO.

	Args:
		path(str): Path to file to setup the print area.
	"""
	v_page = 0
	h_page = 0
	SetPrintArea(path=path, v_page=v_page, h_page=h_page)
	
def SetPrintAreaAPage(path:str) -> None:
	"""Setup the print area into 1 page.
	
	Setup the print area of all sheets in a file specified by argument into a page.

	Args:
		path(str): Path to file to setup the print area.
	
	"""
	v_page = 1
	h_page = 1
	SetPrintArea(path, v_page=v_page, h_page=h_page)

def SetPrintArea(path:str, v_page:int, h_page:int) -> None:
	"""Setup print area in a file.
	
	Setup the vertical and horizontal print range of all sheets in a file, specified by the argument, 
	to the specified values by the argument.

	Args:
		path(str): Path to file to setup the print range.
	
	"""
	wb = openpyxl.load_workbook(path)
	for sheet_item in wb.worksheets:
		_SetupPrintAreaInASheet(sheet=sheet_item, v_page=v_page, h_page=h_page)

	wb.save(path)
	wb.close()

def _SetupPrintAreaInASheet(sheet:worksheet, v_page:int, h_page:int) -> None:
	"""Set print area of a sheet.

	Setup the vertical and horizontal print range of a sheet to the specified values by the argument.

	Args:
		sheet(worksheet): Worksheet to setup the print area.
		v_page(int): The number of page in vertical direction.
		h_page(int): The nubmer of apge in horizontal direction.

	"""
	page_setup = sheet.page_setup
	page_setup.fitToHeight = v_page
	page_setup.fitToWidth = h_page

	sheet_property = sheet.sheet_properties
	sheet_property.pageSetUpPr.fitToPage = True

if '__main__' == __name__:
	path = sys.argv[1]
	SetPrintAreaAPage(path=path)
