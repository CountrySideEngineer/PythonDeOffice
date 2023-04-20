import abc

class ExcelPageConfigFormatter(metaclass=abc.ABCMeta):
	def __init__(self, contest : list):
		self.Content = list

	@abc.abstractclassmethod
	def Write(self, path : str) -> None:
		raise NotImplementedError()

	@abc.abstractclassmethod
	def Read(self, path : str) -> list:
		raise NotImplementedError()
	