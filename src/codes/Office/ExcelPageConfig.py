import abc
from OfficePageConfig import IOfficePageConfig
from ExcelPageConfigCsvFormatter import ExcelPageConfigCsvFormatter as csv_formatter

class ExcelPageConfig(IOfficePageConfig):
	def __init__(self):
		super().__init__

		self.config = list
		self.formatter = csv_formatter()

	@abc.abstractclassmethod
	def Import(self, path : str) -> None:
		raise NotImplementedError()

	@abc.abstractclassmethod
	def Export(self, path : str) -> None:
		raise NotImplementedError()





