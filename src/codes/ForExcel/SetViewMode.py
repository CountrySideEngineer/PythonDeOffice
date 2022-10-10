import sys
import openpyxl

def SetViewMode(path:str, mode:str):
	"""Set sheet mode.

	Set sheet in a file specified by argument, path, into a mode specified by argument, mode.

	Args:
		path(str): Path to file to change, set mode.
		mode(str): Mode to set
	
	"""
	wb = openpyxl.load_workbook(path)
	for sheet in wb.worksheets:
		SetViewModeOfSheet(sheet, mode=mode)
	
	wb.save(path)
	wb.close()

def SetViewModeOfSheet(sheet, mode:str) -> None:
	"""Set sheet mode.

	Set sheet into a mode specified by argument, mode.

	Args:
		sheet(Worksheet): Worksheet to change mode.
		mode(str): Mode to set
	
	"""
	sv = sheet.sheet_view
	sv.view = mode

def SetViewModeNormal(path:str):
	"""Set view mode 'normal'
	
	Set all sheets in a file specified by argument to 'normal' mode.

	Args:
		path(str): Path to file to set all sheets to 'normal' mode
	
	"""
	mode = 'normal'
	SetViewMode(path=path, mode=mode)

def SetViewModePageBreak(path:str):
	"""Set view mode 'break preview'
	
	Set all sheets in a file specified by argument to 'break preview' mode.

	Args:
		path(str): Path to file to set all sheets to 'break preview' mode
	
	"""
	mode = 'pageBreakPreview'
	SetViewMode(path=path, mode=mode)

def SetViewModeLayout(path:str) -> None:
	"""Set view mode 'layout'
	
	Set all sheets in a file specified by argument to 'layuout' mode.

	Args:
		path(str): Path to file to set all sheets to 'layout' mode
	
	"""
	mode = 'pageLayout'
	SetViewMode(path=path, mode=mode)
	
if '__main__' == __name__:
	file_path = sys.argv[1]
	mode = int(sys.argv[2])

	if 0 == mode:
		SetViewModeNormal(path=file_path)
	elif 1 == mode:
		SetViewModePageBreak(path=file_path)
	elif 2 == mode:
		SetViewModeLayout(path=file_path)
	else:
		print('Invalid mode.')

