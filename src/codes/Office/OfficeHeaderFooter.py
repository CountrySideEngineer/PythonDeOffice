import abc
import HeaderFooterItem

class OfficeHeaderFooter():
	def __init__(self):
		"""Constructor.
		
		"""
		self.name = ""
		self.headers = []
		self.footers = []
		
	def ToString(self):
		print('name : ', self.name)
		print('num of headers : ', len(self.headers))
		for header_item in self.headers:
			header_item.ToString()

		print('num of footers : ', len(self.footers))
		for footer_item in self.footers:
			footer_item.ToString()


