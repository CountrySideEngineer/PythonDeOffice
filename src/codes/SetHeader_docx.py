import sys
import docx
from docx import Document

def SetHeaderRight(path:str, headers:list) -> None:
	document = Document(path)
	sections = document.sections
	section_header = sections[0].header

	RemoveHeader(header_section=section_header)

	for header_item in headers:
		section_header.add_paragraph(header_item)

	document.save(path)

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

	SetHeaderRight(path, headers=headers)
