from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart

def GetFooter(footer_part:_HeaderFooterPart) -> str:
	"""Get footer content.

	Args:
		footer_part(_HeaderFooterPart): _HeaderFooterPart object to get from footer.
	
	
	"""
	footer_text = footer_part.text

	return footer_text
