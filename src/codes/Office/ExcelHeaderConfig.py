from ExcelPageConfig import ExcelPageConfig

class ExcelHeaderConfig(ExcelPageConfig):
	def __init__(self):
		"""Constructor.
		"""
		super().__init__()

	def Export(self, path : str) -> None:
		file = open(path, mode = 'w')

		file.close()



	def Import(self, path : str) -> None:



	