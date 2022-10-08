from openpyxl import worksheet
from openpyxl.worksheet.header_footer import _HeaderFooterPart
import JointText as jt

def SetFooter(footer_part:_HeaderFooterPart, footers:list) -> None:
	"""Set header.

	Args:
		footer_part(_HeaderFooterPart): _HeaderFooterPart object to set the headers.
		headetrs(list): Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.
	"""
	fooder_text = jt.JoinText(headers=footers)
	footer_part.text = fooder_text

