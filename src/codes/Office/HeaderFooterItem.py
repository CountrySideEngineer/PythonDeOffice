import abc

class HeaderFooterItem():
	def __init__(self, item : str = '') -> None:
		"""Default constructor.

		"""
		self.item = item		# string object as header or foooter content.

	def ToString(self) :
		print('item = ', self.item)

