import sys
import os
import subprocess
import openpyxl as xl
import xlrd as x


# verifies that a given filename has .xlsx extension
# if it does not, an error is thrown
def check_extension(filename):
	extension_flag = 0
	extension = ""
	good_extension = "xlsx"

	for char in filename:
		if extension_flag==0:
			if char==".":
				extension_flag = 1
		else:
			extension+=char

	if extension!=good_extension:
		sys.exit("must pass in .xlsx files")


# verifies that file exists
# if it does not, an error is thrown
def check_existence(filename):
	if not os.path.isfile(filename):
		sys.exit("file must exist")

def convert_xls(filename):
	xls_filename = filename[0:-4]+"xls"
	subprocess.call(["ssconvert", "--export-type=Gnumeric_Excel:excel_dsf", filename, xls_filename])
	return(xls_filename)


#ssconvert --export-type="Gnumeric_Excel:excel_dsf"

def parse_xlsx(filename):
	book = x.open_workbook(filename, formatting_info=True)
	for sheet in book.sheets():
		sheet_name = sheet.name.lower()
		print(sheet_name)
		if "debrief" in sheet_name:
			debrief_sheet = sheet_name

	if not 'debrief_sheet' in locals():
		sys.exit("no debrief sheet found in .xlsx file")
	else:
		print("found debrief sheet")



def main():
	print("checking command line arguments")
	if len(sys.argv)!=2:
		sys.exit("usage: python parse.py filename.xlsx")
	else:
		print("verifying file extension")
		check_extension(sys.argv[1])
		print("verifying existence")
		check_existence(sys.argv[1])
		filename = sys.argv[1]
		print("converting to xls")
		xls_filename = convert_xls(filename)
		print("parsing xls")
		parse_xlsx(xls_filename)


	

if __name__ == '__main__':
	main()