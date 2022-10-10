from csv import excel
import sys
import openpyxl

default_zoom_scale = 100

def Magnify100Percent(excel_file_path:str) -> None:
	"""Set magnification of sheets to 100 percent.

	Set the magnification in all sheets in a file specified by argument to 100 percent.

	Args:
		excel_file_path (str): Path to excel file to set the mmaginification to 100 percent.
	
	"""
	try:
		wb = openpyxl.load_workbook(excel_file_path)

		for ws in wb.worksheets:
			SetMagnification100Per(ws)

		wb.save(excel_file_path)
		wb.close()
	except Exception:
		print(Exception.message)

def SetMagnification100Per(target_sheet) -> None:
	"""Set magnification of sheet to 100 percent

	Set maginication of a sheet specified by argument to 100 percent.

	Args:
		target_sheet(Worksheet): Sheet to handle in the method.

	"""
	sv = target_sheet.sheet_view
	sv.zoomScale = default_zoom_scale
	sv.soomScalNormal = default_zoom_scale

if '__main__' == __name__:
	file_path = sys.argv[1]
	Magnify100Percent(excel_file_path=file_path)