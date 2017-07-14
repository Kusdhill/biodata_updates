import sys
import os
import subprocess
import openpyxl as xl
import xlrd as x
import xlutils as xu

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

# converts xlsx file to xls file since xlrd will take xls formatting
def convert_xls(filename):
	xls_filename = filename[0:-4]+"xls"
	subprocess.call(["ssconvert", "--export-type=Gnumeric_Excel:excel_dsf",
		filename, xls_filename])
	return(xls_filename)

# look through xls file, find cells with background color
def parse_xls(filename):
	book = x.open_workbook(filename, formatting_info=True)
	for sheet in book.sheets():
		sheet_name = sheet.name.lower()
		if "debrief" in sheet_name:
			debrief_sheet = sheet_name
			break

	if not 'debrief_sheet' in locals():
		sys.exit("no debrief sheet found in .xlsx file")
	else:
		print("found debrief sheet")
		print(sheet.name)

	rows, cols = sheet.nrows, sheet.ncols
	print "Number of rows: %s   Number of cols: %s" % (rows, cols)
	for row in range(rows):
		for col in range(cols):
			cell = sheet.cell(row,col)
			#print(cell.value)

			xfx = sheet.cell_xf_index(row, col)
			xf = book.xf_list[xfx]
			bgx = xf.background.pattern_colour_index
			rgb = book.colour_map[bgx]
			if rgb!=(0,0,0):
				print("in row "+str(row+1)+" column "+str(col)+" : "+str(rgb))

# remove converted xls file
def clean_files(xls_filename):
	os.remove(xls_filename)

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
		parse_xls(xls_filename)
		print("cleaning files")
		clean_files(xls_filename)

if __name__ == '__main__':
	main()