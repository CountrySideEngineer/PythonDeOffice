import sys

from ExcelFooterConfig import ExcelFooterConfig as FooterConfig
from ExcelFileFooter import ExcelFileFooter as FileFooter

if len(sys.argv) < 2:
	print('invalid argument.')

	quit()

csv_file = sys.argv[1]
excel_file = sys.argv[2]

# Read footer from excel files
file_footer = FileFooter(excel_file)
footer_contents = file_footer.Read()

# Export into csv file
footer_config = FooterConfig()
footer_config.config = footer_contents
footer_config.Export(csv_file)
