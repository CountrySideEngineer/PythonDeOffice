import abc

class IOfficePageConfig(metaclass=abc.ABCMeta):

	@abc.abstractclassmethod
	def Import(self, path : str) -> None:
		raise NotImplementedError()

	@abc.abstractclassmethod
	def Export(self, path : str) -> None:
		raise NotImplementedError()
		