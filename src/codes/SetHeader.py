import sys
import openpyxl
from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart

def HeaderFuncPointer(path:str, headers:list, header_func) -> None:
	"""Set header of excel file.
	
	Set header of all sheets in a excel file.

	Args:
		path(str): Path to file to set header.
		headers(list): Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
		header_func: Function used to set header.	
	
	"""
	wb = openpyxl.load_workbook(path)

	sheets = wb.worksheets
	for sheet in sheets:
		header_func(sheet=sheet, headers=headers)

	wb.save(path)
	wb.close()

def SetHeaderRight(path:str, headers:list) -> None:
	"""Set header of right.

	Setup right header of all sheets in a excel file.

	Args:
		path(str): Path to file to set header.
		headers(list): 	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_func = SetPageHeaderRight
	HeaderFuncPointer(path=path, headers=headers, header_func=header_func)

def SetPageHeaderRight(sheet:worksheet, headers:list) -> None:
	"""Set header of right.

	Setup right header of all sheets in a excel file.

	Args:
		sheet(worksheet): Worksheet object to set header.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_part = sheet.oddHeader.right
	SetHeader(header_part=header_part, headers=headers)

def SetHeaderLeft(path:str, headers:list) -> None:
	"""Set header of left.

	Setup left header of all sheets in a excel file.

	Args:
		path(str): Path to file to set header.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_func = SetPageHeaderLeft
	HeaderFuncPointer(path=path, headers=headers, header_func=header_func)

def SetPageHeaderLeft(sheet:worksheet, headers:list) -> None:
	"""Set header of left.

	Setup left header of all sheets in a excel file.

	Args:
		sheet(worksheet): Worksheet object to set header.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_part = sheet.oddHeader.left
	SetHeader(header_part=header_part, headers=headers)

def SetHeaderCenter(path:str, headers:list) -> None:
	"""Set header of center.

	Setup center header of all sheets in a excel file.

	Args:
		path(str): Path to file to set header.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_func = SetPageHeaderCenter
	HeaderFuncPointer(path=path, headers=headers, header_func=header_func)

def SetPageHeaderCenter(sheet:worksheet, headers:list) -> None:
	"""Set header of center.

	Setup center header of all sheets in a excel file.

	Args:
		sheet(worksheet): Worksheet object to set header.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	header_part = sheet.oddHeader.left
	SetHeader(header_part=header_part, headers=headers)

def JoinHeaderText(headers:list) -> str:
	"""Join strings by change line code.
	
	Join strings by change line code as headers

	Args:
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.

	Returns:
		String to be set into header.
	"""
	header_text = ''
	is_top = True
	for header_item in headers:
		if False == is_top:
			header_text += '\n'
		header_text += header_item
		is_top = False

	return header_text

def SetHeader(header_part:_HeaderFooterPart, headers:list) -> None:
	"""Set header.

	Args:
		header_part(_HeaderFooterPart): _HeaderFooterPart object to set the headers.
		headers(list):	Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.

	"""
	headers_text = JoinHeaderText(headers=headers)
	header_part.text = headers_text

if '__main__' == __name__:
	path = sys.argv[1]
	argc = len(sys.argv)
	headers = sys.argv[1:argc]
	SetHeaderRight(path=path, headers=headers)
