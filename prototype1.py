# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
from openpyxl.formula import Tokenizer
import psycopg

debug = True

# input_xlsx_fn = './xlsx/sample-spreadsheet1.xlsx'
input_xlsx_fn = './xlsx/template-spreadsheet.xlsx'

# Pseudo-constants for coordinates of various fields in 'Overview' sheet.
date_coords = 'C2'
temperature_coords = 'C3'
sky_coords = 'C4'
#
bp_loc_id_coords = 'H5'	 # This is a hidden cell in the 'Overview' sheet
loc_desc_coords = 'D5'
loc_desc_other_coords = 'F5'
fac_name_coords = 'D6'
fac_name_other_coords = 'F7'
#
street_1_coords = 'D7'
street_1_dir_coords = 'D8'
street_2_coords = 'D9'
street_2_dir_coords = 'D10'
#
loc_type_coords = 'D11'
muni_coords = 'D12'
muni_other_coords = 'F12'
#
comments_coords = 'J4'

# Pseudo-constants for 'indices' of count-types in the count sheets.
bike_col = 'B'
ped_col = 'C'
child_col = 'D'
jogger_col = 'E'
skater_col = 'F'
wheelchair_col = 'G'
other_col = 'H'

# Pseudo-constants for columns of interest in the 'Columns' sheet
count_loc_desc_col = 'D'
count_loc_id_col = 'E'

# Lists for ranges or row numbers with data in the count sheets.
# Note that count sheet 1 has fewer rows than the other four count sheets.
sheet_1_rows = range(2,12) # i.e., 2 to 11
sheet_2_rows = sheet_3_rows = sheet_4_rows = sheet_5_rows = range(2, 14) # i.e., 2 to 13


wb = None
overview_sheet = None
count_sheet_1 = None
count_sheet_3 = None
count_sheet_4 = None
count_sheet_5 = None
columns_sheet = None  # Sheet containing lookup tables

lookup_table = None	  # Table to map countloc description to countloc id

def initialize(input_fn):
	global wb, overview_sheet, count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4, count_sheet_5, columns_sheet
	wb = openpyxl.load_workbook(filename = input_fn)
	overview_sheet = wb['Overview']
	count_sheet_1 = wb['630-845 AM']
	count_sheet_2 = wb['900-1145 AM']
	count_sheet_3 = wb['1200-245 PM']
	count_sheet_4 = wb['300-545 PM']
	count_sheet_5 = wb['600-845 PM']
	columns_sheet = wb['Columns']
# end_def


# read_overview_sheet: read and parse data from 'Overview' sheet
# parameter: 'lut' - lookup table to map countloc description to id,
#					 bult by read_columns_sheet().
#
def read_overview_sheet(lut):
	global overview_sheet, debug
	
	loc_desc = overview_sheet[loc_desc_coords].value
	
	# Get bp_loc_id from lookup table (lut)
	#
	bp_loc_id = 99999 # Error value
	for row in lut:
		if row['desc'] == loc_desc:
			bp_loc_id = row['id']
			break
		#
	#
	if debug:
		print('bp_loc_id = ' + str(bp_loc_id))
	
	date_raw = overview_sheet[date_coords].value
	if date_raw == None:
		# Temp workaround, for now
		date_raw = '10/23/2023'
	#
	
	# Not sure what to do with the following:
	#
	# if loc_desc == None:
	#	loc_desc = ''
	# elif loc_desc == 'Other':
	#	loc_desc = overview_sheet[loc_desc_other_coords].value
	#
	
	loc_type = overview_sheet[loc_type_coords].value
	
	muni = overview_sheet[muni_coords].value
	if muni == 'Other':
		muni = overview_sheet[muni_other_coords].value
	#	 
	
	fac_name = overview_sheet[fac_name_coords].value
	if fac_name == 'Other':
		fac_name = overview_sheet[fac_name_other_coords].value
	#
	
	street_1 = overview_sheet[street_1_coords].value
	if street_1 == None:
		street_1 = ''
	street_1_dir = overview_sheet[street_1_dir_coords].value
	street_2 = overview_sheet[street_2_coords].value
	if street_2 == None:
		street_2 = ''
	street_2_dir = overview_sheet[street_2_dir_coords].value

	temperature = overview_sheet[temperature_coords].value
	if temperature == None:
		temperature = ''
	#
	
	sky = overview_sheet[sky_coords].value
	if sky == 'Sunny':
		sky = 1
	elif sky == 'Partly Cloudy':
		sky = 2
	elif sky == 'Overcast':
		sky = 3
	elif sky == 'Precipitation':
		sky = 4
	elif sky == 'No Data':
		sky = 99
	else:
		sky = 99
	# end_if
	
	comments = overview_sheet[comments_coords].value
	if comments == None:
		comments = ''
	#
	
	if debug:
		print('bp_loc_id = ' + str(bp_loc_id))
		print('date	 = ' + str(date_raw))
		print('location description = ' + loc_desc)
		print('location type = ' + loc_type)
		print('municipality = ' + muni)
		print('facility name = ' + fac_name)
		print('street 1 = ' + street_1)
		print('street 1 direction = ' + street_1_dir)
		print('street 2 = ' + street_2_dir)
		print('street 2 direction = ' + street_2_dir)
		print('temperature = ' + str(temperature))
		print('sky = ' + str(sky))
		print('comments = '	 + comments)
	# end_if 
		
	# *** TODO: Clean up raw date
	date_cooked = date_raw
	
	# Assemble return value: dict of information harvested from overview table
	retval = { 'bp_loc_id' : bp_loc_id, 'date' : date_cooked, 
			   'street_1' : street_1, 'street_1_dir' : street_1_dir,
			   'street_2' : street_2, 'street_2_dir' : street_2_dir,   
			   'temperature' : temperature, 'sky' : sky,
			   'comments' : comments }
	return retval
# end_def: read_overview_sheet

# read_columns_sheet - read data from 'Columns' sheet;
# specifically create "count location description-to-bp_loc_id" lookup table
# from the contents of columns D and E in this sheet.
#
def read_columns_sheet():
	global columns_sheet, count_loc_desc_col, count_loc_id_col
	
	row_ix = 2
	ix = count_loc_desc_col + str(row_ix)
	val = columns_sheet[ix].value
	if debug:
		print('Value of cell ' + ix + ' is: ' + val)
	while val != None:
		row_ix = row_ix + 1
		ix = count_loc_desc_col + str(row_ix)
		val = columns_sheet[ix].value
	# 
	if debug:
		print('Last row_ix was: ' + str(row_ix))
		
	# Note: row_ix is one beyond index of the last row w/ real data
	lookup_table = []
	for row in range(2, row_ix):
		desc_ix = count_loc_desc_col + str(row)
		id_ix = count_loc_id_col + str(row)
		temp = { 'desc' : columns_sheet[desc_ix].value, 'id' : columns_sheet[id_ix].value }
		lookup_table.append(temp)
	#
	return lookup_table
# end_def read_columns_sheet


# read_count_sheet: Read data from _one_ count sheet.
#
# Parameters:
#	  count_sheet - workbook count sheet to be read
#	  rows - range of rows to be read in count sheet
#
def read_count_sheet(count_sheet, rows):
	global debug
	
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
	if debug:
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
	if debug:
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
	if debug:
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
	if debug:
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
	if debug:
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
	if debug:
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
	if debug:
		print('Other counts:')
		for c in other_temp:
			print(c)
		#
	#
	
	# Assemble return value: dict of list of each count type
	retval = { 'bike' : bike_temp, 'ped': ped_temp, 'child' : child_temp,
			   'jogger' : jogger_temp, 'skater' : skater_temp,
			   'wheelchair' : wheelchair_temp, 'other' : other_temp }
	return retval
# end_def: read_count sheet

# read_count_sheets - read data from all count sheets;
# a 'driver routine' that calls read_count_sheet for 
# each of the 5 'count' sheets'
#
def read_count_sheets():
	global count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4, count_sheet_5
	global sheet_1_rows, sheet_2_rows, sheet_3_rows, sheet_4_rows, sheet_5_rows
	s1_data = read_count_sheet(count_sheet_1, sheet_1_rows)
	s2_data = read_count_sheet(count_sheet_2, sheet_2_rows)
	s3_data = read_count_sheet(count_sheet_3, sheet_3_rows)
	s4_data = read_count_sheet(count_sheet_4, sheet_4_rows)
	s5_data = read_count_sheet(count_sheet_5, sheet_5_rows)
	# Assemble count data from all sheets
	bike_data = s1_data['bike'] + s2_data['bike'] + s3_data['bike'] + s4_data['bike'] + s5_data['bike']
	ped_data = s1_data['ped']+ s2_data['ped'] + s3_data['ped'] + s4_data['ped'] + s5_data['ped']
	child_data = s1_data['child']+ s2_data['child'] + s3_data['child'] + s4_data['child'] + s5_data['child']
	jogger_data = s1_data['jogger']+ s2_data['jogger'] + s3_data['jogger'] + s4_data['jogger'] + s5_data['jogger']
	skater_data = s1_data['skater']+ s2_data['skater'] + s3_data['skater'] + s4_data['skater'] + s5_data['skater']
	wheelchair_data = s1_data['skater']+ s2_data['skater'] + s3_data['skater'] + s4_data['skater'] + s5_data['skater']
	other_data = s1_data['other']+ s2_data['other'] + s3_data['other'] + s4_data['other'] + s5_data['other']	 
	 # Assemble return value
	retval = { 'bike' : bike_data, 'ped' : ped_data, 'child' : child_data,
			   'jogger' : jogger_data, 'skater' : skater_data,
			   'wheelchair' : wheelchair_data, 'other' : other_data }
	return retval	
# end_def: read_count_sheets

# Test uber-driver routine:
def test_driver():
	initialize(input_xlsx_fn)
	lut = read_columns_sheet() # build lookup table
	overview_data = read_overview_sheet(lut)
	count_data = read_count_sheets()
	# Here: Have all info needed to assemble and run SQL INSERT query
# end_def: test_driver

# Test driver for only reading 'Overview' sheet
def test_driver_overview():
	initialize(input_xlsx_fn)
	overview_data = read_overview_sheet()
# end_def

# Test driver for reading 'Columns' sheet and constructing lookup table
def test_driver_columns_sheet():
	initialize(input_xlsx_fn)
	lut = read_columns_sheet()
	print('Dump of LUT:')
	for row in lut:
		print(str(row['id']))
	#
# end_def