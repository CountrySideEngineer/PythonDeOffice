import sys

from ExcelFooterConfig import ExcelFooterConfig as FooterConfig
from ExcelFileFooter import ExcelFileFooter as FileFooter

if len(sys.argv) < 2:
	print('invalid argument.')

	quit()

csv_file = sys.argv[1]
excel_file = sys.argv[2]

# Import Excel footer contents
footer_config = FooterConfig()
footer_config.config = []
footer_config.Import(csv_file)

# Export header contents into excel
file_footer = FileFooter(excel_file)
file_footer.WriteAll(footer_config.config)

