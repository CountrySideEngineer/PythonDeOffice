from csv import excel
import sys
import openpyxl

default_zoom_scale = 100

def Magnify100Percent(excel_file_path):
	try:
		wb = openpyxl.load_workbook(excel_file_path)

		for ws in wb.worksheets:
			SetMagnification100Per(ws)

		wb.save(excel_file_path)
	except Exception:
		print(Exception.message)

def SetMagnification100Per(target_sheet):
	sv = target_sheet.sheet_view
	sv.zoomScale = default_zoom_scale
	sv.soomScalNormal = default_zoom_scale

if '__main__' == __name__:
	file_path = sys.argv[1]
	Magnify100Percent(excel_file_path=file_path)