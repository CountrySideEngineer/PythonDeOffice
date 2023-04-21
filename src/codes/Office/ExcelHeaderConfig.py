from ExcelPageConfig import ExcelPageConfig
from OfficeHeaderFooter import OfficeHeaderFooter

class ExcelHeaderConfig(ExcelPageConfig):
	def __init__(self):
		"""Constructor.
		"""
		super().__init__()

	def Export(self, path : str) -> None:
		"""Export excel page config to a file.

		Exports excel page config as OfficeHeaderFooter object into a file.

		Args:
			path(str) : Path to file to export.
		
		"""
		config_as_lsit = self.ConfigToList()
		self.formatter.content = config_as_lsit
		self.formatter.Write(path=path)

	def ConfigToList(self) -> list:
		"""Convert excel page configuration into list.

		Converts excel page configuration as OfficeHeaderFooter object.

		Returns:
			Config parameters as list
		"""
		headers = []
		for item in self.config:
			name = item.name
			header_item = []
			header_item.append(name)
			for header in item.headers:
				header_item.append(header)

			headers.append(header_item)

		return list(headers)

	def Import(self, path : str) -> None:
		"""Imports excel page cofig as OfficeHeaderFooter object from file.

		Args:
			path(str) : Path to file to import from.
		"""
		self.formatter.Read(path)
		self.ListToConfig(self.formatter.content)

	def ListToConfig(self, item_list : list) -> None:
		self.config = []
		for item in item_list:
			config = OfficeHeaderFooter()
			config.name = item[0]
			config.headers.append(item[1])
			config.headers.append(item[2])
			config.headers.append(item[3])

			self.config.append(config)

if __name__ == '__main__':
	item1 = OfficeHeaderFooter()
	item1.name = 'sheet1'
	item1.headers = ['item1-11', 'item1-12', 'item1-13']

	item2 = OfficeHeaderFooter()
	item2.name = 'sheet2'
	item2.headers = ['item2-11', 'item2-12', 'item2-13']

	contents = []
	contents.append(item1)
	contents.append(item2)

	config = ExcelHeaderConfig()
	config.config = contents
	path = r'E:\development\PythonDeOffice\src\samples\ExcelHeaderConfig_csv_sample.csv'
	config.Export(path)

	config.config = []
	config.Import(path)

	for config_item in config.config:
		print(config_item.name, ',', ' , '.join(config_item.headers))
