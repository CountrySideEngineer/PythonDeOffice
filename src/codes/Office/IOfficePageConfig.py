import abc

class IOfficePageConfig(metaclass=abc.ABCMeta):
	def __init__(self):
		"""Default constrcutor
		"""

	@abc.abstractmethod
	def Import(self, path : str) -> None:
		raise NotImplementedError()

	@abc.abstractmethod
	def Export(self, path : str) -> None:
		raise NotImplementedError()
