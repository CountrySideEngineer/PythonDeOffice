import openpyxl
from openpyxl import Workbook
from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart
from ExcelFile import ExcelFile
from OfficeHeaderFooter import OfficeHeaderFooter
from HeaderFooterItem import HeaderFooterItem

class ExcelFileFooter(ExcelFile):
	def __init__(self, path : str = ""):
		"""Constructor.
		
		"""
		super().__init__(path)

	def AppendItem(self, dst : OfficeHeaderFooter, item : HeaderFooterItem) -> None:
		"""Append footer object as OffceHeaderFooter object to 
		
		"""
		dst.footers.append(item)

	def GetLeftPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		part = sheet.oddFooter.left
		return part

	def GetCenterPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		part = sheet.oddFooter.center
		return part

	def GetRightPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		part = sheet.oddFooter.right
		return part

if __name__ == '__main__':
	path = r'E:\development\PythonDeOffice\src\samples\ActiveLeftSheet_sample.xlsx'
	footer = ExcelFileFooter(path)

	footers = footer.Read()