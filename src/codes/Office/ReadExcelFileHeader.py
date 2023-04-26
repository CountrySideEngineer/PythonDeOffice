import sys

from ExcelHeaderConfig import ExcelHeaderConfig as HeaderConfig
from ExcelFileHeader import ExcelFileHeader as FileHeader

if len(sys.argv) < 2:
	print('invalid argument.')

	quit()

csv_file = sys.argv[1]
excel_file = sys.argv[2]

# Read header from excel files
file_header = FileHeader(excel_file)
header_contents = file_header.Read()

# Export into csv file
header_config = HeaderConfig()
header_config.config = header_contents
header_config.Export(csv_file)
