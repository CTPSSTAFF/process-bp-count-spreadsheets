# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
import psycopg

debug = True

# Pseudo-constants for coordinates of various fields in 'Overview' sheet.
date_coords = 'B2'
muni_coords = 'B3'
loc_type_coords = 'B4'
loc_desc_coords = 'B5'
fac_name_coords = 'B6'
from_st_coords = 'B7'
from_st_dir_coords = 'B8'
to_st_coords = 'B9'
to_st_dir_coords = 'B10'
temp_coords = 'B11'
sky_coords = 'B12'
comments_coords = 'A15'

# Pseudo-constants for 'indices' of count-types in the count sheets.
bike_col = 'B'
ped_col = 'C'
child_col = 'D'
jogger_col = 'E'
skater_col = 'F'
wheelchair_col = 'G'
other_col = 'H'

# Lists into which count data from the count sheets in the workbook will be accumulated
bike_data = []
ped_data = []
child_data = []
jogger_data = []
skater_data	 = []
wheelchair_data = []
other_data = []

# Lists for ranges or row numbers with data in the count sheets.
# Note that count sheet 1 has fewer rows than the other four count sheets.
sheet_1_rows = range(2,12) # i.e., 2 to 11
sheet_2_rows = sheet_3_rows = sheet_4_rows = sheet_5_rows = range(2, 14) # i.e., 2 to 13

input_xlsx_fn = './xlsx/sample-spreadsheet1.xlsx'

wb = None
overview_sheet = None
count_sheet_1 = None
count_sheet_3 = None
count_sheet_4 = None
count_sheet_5 = None

def initialize(input_fn):
	global wb, overview_sheet, count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4, count_sheet_5
	wb = openpyxl.load_workbook(filename = input_fn)
	overview_sheet = wb['Overview']
	count_sheet_1 = wb['630-845 AM']
	count_sheet_2 = wb['900-1145 AM']
	count_sheet_3 = wb['1200-245 PM']
	count_sheet_4 = wb['300-545 PM']
	count_sheet_5 = wb['600-845 PM']
# end_def


def read_overview_tab():
	global overview_sheet, debug
	date_raw = overview_sheet[date_coords].value
	if debug:
		print('date ' + str(date_raw))
	
	muni = overview_sheet[muni_coords].value
	if debug:
		print('municipality = ' + muni)
	
	loc_type = overview_sheet[loc_type_coords].value
	if debug:
		print('location type = ' + loc_type)
	
	loc_desc = overview_sheet[loc_desc_coords].value
	if debug:
		print('location description = ' + loc_desc)
	
	fac_name = overview_sheet[fac_name_coords].value
	if debug:
		print('facility name = ' + fac_name)
	
	from_st = overview_sheet[from_st_coords].value
	if debug:
		print('from street = ' + from_st)
	
	from_st_dir = overview_sheet[from_st_dir_coords].value
	if debug:
		print('from street direction = ' + from_st_dir)
	
	to_st = overview_sheet[to_st_coords].value
	if debug:
		print('from street = ' + to_st)
	
	to_st_dir = overview_sheet[to_st_dir_coords].value
	if debug:
		print('from street direction = ' + to_st_dir)
	
	temp = overview_sheet[temp_coords].value
	if debug:
		print('temperature = ' + str(temp))
	
	sky = overview_sheet[sky_coords].value
	if debug:
		print('sky = ' + sky)
	
	comments = overview_sheet[comments_coords].value
	if debug:
		print('comments = '	 + comments)
# end_def: read_overview_tab

# Read data from one count sheet.
# Parameters:
#	  count_sheet - workbook count sheet to be read
#	  rows - range of rows to be read in count sheet
#
def read_count_sheet(count_sheet, rows):
	global debug
	global bike_data, ped_data, child_data, jogger_data, skater_data, wheelchair_data, other_data
	
	bike_temp = []
	for row in rows:
		ix = bike_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		bike_temp.append(val)
	#
	if (debug == True):
		print('Bike counts:')
		for c in bike_temp:
			print(c)
		#
	#
	
	ped_temp = []
	for row in rows:
		ix = ped_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		ped_temp.append(val)
	#
	if (debug == True):
		print('Ped counts:')
		for c in ped_temp:
			print(c)
		#
	#
	
	child_temp = []
	for row in rows:
		ix = child_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		child_temp.append(val)
	#
	if (debug == True):
		print('Child counts:')
		for c in child_temp:
			print(c)
		#
	#
	
	jogger_temp = []
	for row in rows:
		ix = jogger_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		jogger_temp.append(val)
	#
	if (debug == True):
		print('Jogger counts:')
		for c in jogger_temp:
			print(c)
		#
	#
	
	skater_temp = []
	for row in rows:
		ix = skater_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		skater_temp.append(val)
	#
	if (debug == True):
		print('Skater counts:')
		for c in skater_temp:
			print(c)
		#
	#
	
	wheelchair_temp = []
	for row in rows:
		ix = wheelchair_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		wheelchair_temp.append(val)
	#
	if (debug == True):
		print('Wheelchair counts:')
		for c in wheelchair_temp:
			print(c)
		#
	#
	
	other_temp = []
	for row in rows:
		ix = other_col + str(row)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = 0
		#
		other_temp.append(val)
	#
	if (debug == True):
		print('Other counts:')
		for c in other_temp:
			print(c)
		#
	#
	
	# TBD: Next step(s)
# end_def: read_count sheet

# Driver routine: read data from all count sheets.
def read_count_sheets():
	global count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4, count_sheet_5
	global sheet_1_rows, sheet_2_rows, sheet_3_rows, sheet_4_rows, sheet_5_rows
	pass
	read_count_sheet(count_sheet_1, sheet_1_rows)
# end_def: read_count_sheets

# Test uber-driver routine:
def test_driver():
	initialize(input_xlsx_fn)
	read_count_sheets()
# end_def: test_driver