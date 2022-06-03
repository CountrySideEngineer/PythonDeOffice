from csv import excel
import sys
import openpyxl

def ActivateLeftMostSheet(excel_file_path):
	"""Activate the leftmost sheet in a file.

	Activate the leftmost in a file and inactivate the selected sheet
	when the opening file.
	If the sheet is the leftmost one, there will be no change.
	
	Args:
		excel_file_path(string):	Path to file to chagne.

	Returns:
		None

	"""
	wb = openpyxl.load_workbook(excel_file_path)

	#Inactivate all sheet
	for sheet_item in wb.worksheets:
		sheet_item.sheet_view.tabSelected = False

	#Activate leftmost sheet in a file.
	wb.active = wb.worksheets[0]

	wb.save(excel_file_path)
	wb.close()

if '__main__' == __name__:
	excel_file_path = sys.argv[1]
	ActivateLeftMostSheet(excel_file_path=excel_file_path)
