import abc

class ExcelPageConfigFormatter(metaclass=abc.ABCMeta):
	def __init__(self, contest : list):
		self.Content = list

	@abc.abstractmethod
	def Write(self, path : str) -> None:
		raise NotImplementedError()
	