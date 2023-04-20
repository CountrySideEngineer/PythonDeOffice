import sys
import csv
from ExcelPageConfigFormatter import ExcelPageConfigFormatter

class ExcelPageConfigCsvFormatter(ExcelPageConfigFormatter):
	def __init__(self, content : list = None):
		self.content = content

	def Write(self, path : str) -> None:
		file = open(path, newline='', mode='w')
		csv_writer = csv.writer(file)
		csv_writer.writerows(self.content)
		file.close()

	def Read(self, path : str) -> list:
		rows = []
		file = open(path, mode='r')
		csv_reader = csv.reader(file)
		for row in csv_reader:
			print(', '.join(row))

			rows.append(row)

		file.close()

		return rows

if __name__ == '__main__':
	content = [
		['item11', 'item12', 'item13', 'item14'],
		['item21', 'item22', 'item23', 'item24'],
		['item31', 'item32', 'item33', 'item34'],
		['item41', 'item42', 'item43', 'item44']]
	formatter = ExcelPageConfigCsvFormatter(content)
	path = r'E:\development\PythonDeOffice\src\samples\formatter_csv_sample.csv'
	formatter.Write(path)

	formatter.Read(path)

