import abc
import openpyxl
from openpyxl import Workbook
from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart
import OfficeFile
from OfficeHeaderFooter import OfficeHeaderFooter
from HeaderFooterItem import HeaderFooterItem

class ExcelFile(OfficeFile.IOfficeFile):
	def __init__(self, path : str = ""):
		"""Constructor

		"""
		super().__init__(path = path)

	def Write(self, item : OfficeHeaderFooter) -> None:
		print('Write')

	def WriteAll(self, items : list) -> None:
		print('WriteAll')

	def Read(self) -> list:
		with openpyxl.load_workbook(self.path) as wb:
			items = self.ReadFromBook(wb)

		return items

	def ReadFromBook(self, wb : Workbook) -> list:
		sheet_names = wb.sheetnames
		header_footers = []
		for sheet_name in sheet_names:
			ws = wb[sheet_name]
			hf = self.ReadFromSheet(ws)
			hf.name = sheet_name
			header_footers.append(hf)

		return header_footers

	def ReadFromSheet(self, ws : worksheet) -> OfficeHeaderFooter:
		left_part = self.GetLeftPartFromSheet(sheet = ws)
		left_item = HeaderFooterItem()
		left_item.item = self.ReadFromPart(part = left_part)

		center_part = self.GetCenterPartFromSheet(sheet = ws)
		center_item = HeaderFooterItem()
		center_item.item = self.ReadFromPart(part = center_part)

		right_part = self.GetRightPartFromSheet(sheet = ws)
		right_item = HeaderFooterItem()
		right_item.item = self.ReadFromPart(part = right_part)

		hf = OfficeHeaderFooter()
		self.AppendItem(dst = hf, item = left_item)
		self.AppendItem(dst = hf, item = center_item)
		self.AppendItem(dst = hf, item = right_item)

		return hf

	def ReadFromPart(self, part : _HeaderFooterPart) -> str:
		item = part.text
		return item

	@abc.abstractclassmethod
	def AppendItem(self, dst : OfficeHeaderFooter, item : HeaderFooterItem) -> None:
		raise NotImplementedError()

	@abc.abstractmethod
	def GetLeftPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		raise NotImplementedError()

	@abc.abstractmethod
	def GetCenterPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		raise NotImplementedError()

	@abc.abstractmethod
	def GetRightPartFromSheet(self, sheet : worksheet) -> _HeaderFooterPart:
		raise NotImplementedError()
