import sys

from ExcelHeaderConfig import ExcelHeaderConfig as HeaderConfig
from ExcelFileHeader import ExcelFileHeader as FileHeader

if len(sys.argv) < 2:
	print('invalid argument.')

	quit()

csv_file = sys.argv[1]
excel_file = sys.argv[2]

# Import Excel header contents
header_config = HeaderConfig()
header_config.config = []
header_config.Import(csv_file)

# Export header contents into excel
file_header = FileHeader(excel_file)
file_header.WriteAll(header_config.config)

