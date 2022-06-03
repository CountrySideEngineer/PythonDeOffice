import sys
import openpyxl

def FocusA1Cell(excel_file_path):
	"""Set focus on A1 cell.

	Set focus on A1 cell of each work sheet in a file.

	Args:
		excel_file_path(string):	Path to file to set focus on.

	Returns:
		None
	
	"""
	try:
		wb = openpyxl.load_workbook(excel_file_path)

		for sheet_name in wb.sheetnames:
			ws = wb[sheet_name]
			SetFocusOnA1Cell(ws)

		wb.save(excel_file_path)
		wb.close()
	except Exception:
		print(Exception.message)


def SetFocusOnA1Cell(target_sheet):
	"""Set focus on cell A1

	Set focus on A1 cell on a sheet.

	Args:
		target_sheet(worksheet):	Excel sheet object to set focus on A1 cell.

	Returns:
		None

	"""
	sv = target_sheet.sheet_view
	sv.selection[0].activeCell = 'A1'
	sv.selection[0].sqref = 'A1'
	sv.selection[0].activeCellId = None

if '__main__' == __name__:
	file_path = sys.argv[1]
	FocusA1Cell(file_path)
