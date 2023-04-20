import abc

class HeaderFooterItem():
	def __init__(self) -> None:
		"""Default constructor.

		"""
		self.item = ""		# string object as header or foooter content.

	def ToString(self) :
		print('item = ', self.item)

