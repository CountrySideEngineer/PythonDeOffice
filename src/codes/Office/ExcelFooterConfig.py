from ExcelPageConfig import ExcelPageConfig
from OfficeHeaderFooter import OfficeHeaderFooter
from HeaderFooterItem import HeaderFooterItem as hfitem

class ExcelFooterConfig(ExcelPageConfig):
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
		footers = []
		for item in self.config:
			name = item.name
			footer_item = []
			footer_item.append(name)
			for footer in item.footers:
				footer_item.append(footer.item)

			footers.append(footer_item)

		return list(footers)

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
			try:
				config.name = item[0]
			except IndexError:
				config.name = ''

			try:
				config.footers.append(item[1])
			except IndexError:
				config.footers.append('')

			try:
				config.footers.append(item[2])
			except IndexError:
				config.footers.append('')

			try:
				config.footers.append(item[3])
			except IndexError:
				config.footers.append('')

			self.config.append(config)

if __name__ == '__main__':
	item1 = OfficeHeaderFooter()
	item1.name = 'sheet1'
	item1.footers = [hfitem('footer_item1-11'), hfitem('footer_item1-12'), hfitem('footer_item1-13')]

	item2 = OfficeHeaderFooter()
	item2.name = 'sheet2'
	item2.footers = [hfitem('footer_item2-11'), hfitem('footer_item2-12'), hfitem('footer_item2-13')]

	contents = []
	contents.append(item1)
	contents.append(item2)

	config = ExcelFooterConfig()
	config.config = contents
	path = r'E:\development\PythonDeOffice\src\samples\ExcelFooterConfig_csv_sample.csv'
	config.Export(path)

	config.config = []
	config.Import(path)

	for config_item in config.config:
		print(config_item.name, ',', ' , '.join(config_item.footers))
