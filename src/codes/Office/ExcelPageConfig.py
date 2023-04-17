import abc
from OfficePageConfig import IOfficePageConfig

class ExcePageConfig(IOfficePageConfig):
	def __init__(self):
		super().__init__

		self.config = list

	@abc.abstractclassmethod
	def Import(self, path : str) -> None:
		raise NotImplementedError()

	@abc.abstractclassmethod
	def Export(self, path : str) -> None:
		raise NotImplementedError()





