import abc
from OfficeHeaderFooter import OfficeHeaderFooter

class IOfficeFile(metaclass=abc.ABCMeta):
	def __init__(self, path = "") : 
		self.path = path

	@abc.abstractmethod
	def Write(self, items : list) -> None:
		raise NotImplementedError()

	@abc.abstractmethod
	def WriteAll(self, item : OfficeHeaderFooter) -> None:
		raise NotImplementedError()

	@abc.abstractmethod
	def Read(self) -> list:
		raise NotImplementedError()
