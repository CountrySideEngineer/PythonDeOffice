import sys
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

def SetHeaderLeft(path:str, headers:list) -> None:
	align_func = SetHeaderAlignmentLeft
	HeaderFuncPointer(path=path, headers=headers, align_func=align_func)

def SetHeaderCenter(path:str, headers:list) -> None:
	align_func = SetHeaderAlignmentCenter
	HeaderFuncPointer(path=path, headers=headers, align_func=align_func)

def SetHeaderRight(path:str, headers:list) -> None:
	align_func = SetHeaderAlignmentRight
	HeaderFuncPointer(path=path, headers=headers, align_func=align_func)

def HeaderFuncPointer(path:str, headers:list, align_func) -> None:
	document = Document(path)
	sections = document.sections
	section_header = sections[0].header

	RemoveHeader(header_section=section_header)

	for header_item in headers:
		new_pargraph = section_header.add_paragraph(header_item)
		align_func(new_pargraph)

	document.save(path)

def SetHeaderAlignmentRight(paragrah) -> None:
	paragrah.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def SetHeaderAlignmentLeft(paragrah) -> None:
	paragrah.alignment = WD_ALIGN_PARAGRAPH.LEFT

def SetHeaderAlignmentCenter(paragrah) -> None:
	paragrah.alignment = WD_ALIGN_PARAGRAPH.CENTER

def RemoveHeader(header_section:docx.section.Sections) -> None:
	paragraphs = header_section.paragraphs
	for paragraph in paragraphs:
		element = paragraph._element
		element.getparent().remove(element)
		paragraph._p = paragraph._element = None

if '__main__' == __name__:
	path = sys.argv[1]

	headers = [
		'header_code_011',
		'header_code_044',
		]

	SetHeaderLeft(path, headers=headers)
