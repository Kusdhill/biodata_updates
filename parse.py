import sys
import os
import subprocess
import xlrd as x
from xlutils.copy import copy
import xlwt as xlwt

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
		#print("found debrief sheet")
		print(sheet.name)

	print("finding header location")
	header_location = get_header(sheet, 0)

	print("finding highlights")
	rows, cols = sheet.nrows, sheet.ncols
	rgb_row_col = []
	clean_row_col = []
	for row in range(rows):
		for col in range(cols):
			prev_col = col-1
			cell = sheet.cell(row,col)

			xfx = sheet.cell_xf_index(row, col)
			xf = book.xf_list[xfx]
			bgx = xf.background.pattern_colour_index
			rgb = book.colour_map[bgx]
			if rgb!=(0,0,0):
				#print("highlight in row "+str(row+1)+" column "+str(col)+" : "+str(rgb))
				if row==header_location:
					pass
				else:
					rgb_row_col.append([row,col])
			else:
				clean_row_col.append([row,col])

	#print(rgb_row_col)
	print("writing xls")
	write_xls(filename, sheet, rgb_row_col, clean_row_col, header_location)



def write_xls(filename, sheet, rgb_locations, clean_locations, header_location):

	#print(clean_locations)
	xlwt_workbook = xlwt.Workbook()
	xls_save_filename = filename[0:-4]+"_biodata_updates.xls"
	xlwt_sheet = xlwt_workbook.add_sheet("Debrief", cell_overwrite_ok=False)
	row_pointer = 1

	header_style = xlwt.easyxf('pattern: pattern solid;')
	header_style.pattern.pattern_fore_colour = 150

	style = xlwt.easyxf('pattern: pattern solid;')
	style.pattern.pattern_fore_colour = 50

	for col in range(0,sheet.ncols):
		cell = sheet.cell(header_location,col)
		xlwt_sheet.write(0,col,cell.value,header_style)
	print("written header")

	for i in range(0, len(rgb_locations)):
		row = rgb_locations[i][0]
		prev_row = rgb_locations[i-1][0]
		col = rgb_locations[i][1]
		prev_col = rgb_locations[i-1][1]

		cell = sheet.cell(row,col)
		
		if prev_row==row or row<2:
			#print("passing "+str(row)+" "+str(prev_row)+" "+cell.value+"\n")
			pass
		else:
			#print(str(row)+" "+str(prev_row)+"\n")
			row_pointer+=1
		if cell.value=="":
			pass
		else:
			xlwt_sheet.write(row_pointer,col,cell.value,style)

	rgb_rows = []
	for i in range(0, len(rgb_locations)):
		rgb_row = rgb_locations[i][0]
		rgb_rows.append(rgb_row)
	print(rgb_rows)

	# handles clean col adding based on rgb cols
	clean_rows = {}
	for i in range(0, len(clean_locations)):
		clean_row = clean_locations[i][0]
		clean_col = clean_locations[i][1]
		if not clean_row in clean_rows:
			clean_rows[clean_row] = [clean_col]
		else:
			clean_rows[clean_row].append(clean_col)
	#print(clean_rows)

	# removes rows from clean_rows if they aren't in rgb rows
	# since we only care about rows that have rgb
	print("rows")
	for i in range(0, max(rgb_rows)+1):
		#rgb_row = rgb_locations[i][0]
		
		if i in rgb_rows:
			print(str(i)+" is in rbg, keeping in clean_rows")
		else:
			if i in clean_rows:
				print(str(i)+ " is not in rgb, deleting row")
				del clean_rows[i]
			else:
				pass

    # shave off unnecessary values
	for key in clean_rows.keys():
		if key>max(rgb_rows):
			del clean_rows[key]

	#print(clean_rows)

	print(rgb_locations)
	print(clean_rows.keys())

	row_pointer = 1
	for key in clean_rows:
		for value in clean_rows[key]:
			#print(key,value)
			row = key
			col = value
			cell = sheet.cell(row,col)				
		
			#print("writing "+str(cell.value))
			xlwt_sheet.write(row_pointer,col,cell.value)
		row_pointer+=1


	xlwt_workbook.save(xls_save_filename)


# get header row from sheet
def get_header(sheet, row):
	cols = sheet.ncols
	header_found = False
	debrief_words = ["attended","first name","last name",
		"employer","relationship"]
	column_words = []
	for col in range(cols):
		cell = sheet.cell(row,col)
		value = str(cell.value)
		if type(value) is str:
			column_words.append(value.lower())
		else:
			pass

	for dword in debrief_words:
		for cword in column_words:
			if dword in cword:
				header_found = True

	if header_found:
		#print("header found in row "+str(row))
		return(row)
	else:
		row+=1
		return(get_header(sheet, row))

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
		print("complete!")

if __name__ == '__main__':
	main()