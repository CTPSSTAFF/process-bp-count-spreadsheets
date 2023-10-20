# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
import psycopg

debug = True

# Pseudo-constants for coordinates of various fields in 'Overview' sheet.
date_coords = 'C2'
temperature_coords = 'C3'
sky_coords = 'C4'
#
bp_loc_id_coords = 'H5'	 # This is a hidden cell in the spreadsheet
loc_desc_coords = 'D6'
loc_desc_other_coords = 'F6'
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


def read_overview_sheet():
	global overview_sheet, debug
	
	bp_loc_id = overview_sheet[bp_loc_id_coords].value
	date_raw = overview_sheet[date_coords].value
	
	loc_desc = overview_sheet[loc_desc_coords].value
	if loc_desc == 'Other':
		loc_desc = overview_sheet[loc_desc_other_coords].value
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
	street_1_dir = overview_sheet[street_1_dir_coords].value:
	street_2 = overview_sheet[street_2_coords].value
	street_2_dir = overview_sheet[street_2_dir_coords].value

	temperature = overview_sheet[temperature_coords].value
	sky = overview_sheet[sky_coords].value
	comments = overview_sheet[comments_coords].value
	
	if debug == True:
		print('bp_loc_id = ' + bp_loc_id)
		print('date	 =' + str(date_raw))
		print('location description = ' + loc_desc)
		print('location type = ' + loc_type
		print('municipality = ' + muni)
		print('facility name = ' + fac_name)
		print('street 1 = ' + street_1)
		print('street 1 direction = ' + street_1_dir)
		print('street 2 = ' + street_2_dir)
		print('street 2 direction = ' + street_2_dir)
		print('temperature = ' + str(temperature))
		print('sky = ' + sky)
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

# Read data from one count sheet.
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
	
	# Assemble return value: dict of list of each count type
	retval = { 'bike' : bike_temp, 'ped': ped_temp, 'child' : child_temp,
			   'jogger' : jogger_temp, 'skater' : skater_temp,
			   'wheelchair' : wheelchair_temp, 'other' : other_temp }
	return retval
# end_def: read_count sheet

# Driver routine: read data from all count sheets.
def read_count_tabs():
	global count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4, count_sheet_5
	global sheet_1_rows, sheet_2_rows, sheet_3_rows, sheet_4_rows, sheet_5_rows
	s1_data = read_count_sheet(count_sheet_1, sheet_1_rows)
	s2_data = read_count_sheet(count_sheet_2, sheet_2_rows)
	s3_data = read_count_sheet(count_sheet_3, sheet_3_rows)
	s4_data = read_count_sheet(count_sheet_4, sheet_4_rows)
	s5_data = read_count_sheet(count_sheet_5, sheet_5_rows)
	# Assemble count data from all sheets
	bike_data = s1['bike'] + s2['bike'] + s3['bike'] + s4['bike'] + s5['bike']
	ped_data = s1['ped']+ s2['ped'] + s3['ped'] + s4['ped'] + s5['ped']
	child_data = s1['child']+ s2['child'] + s3['child'] + s4['child'] + s5['child']
	jogger_data = s1['jogger']+ s2['jogger'] + s3['jogger'] + s4['jogger'] + s5['jogger']
	skater_data = s1['skater']+ s2['skater'] + s3['skater'] + s4['skater'] + s5['skater']
	wheelchair_data = s1['skater']+ s2['skater'] + s3['skater'] + s4['skater'] + s5['skater']
	other_data = s1['other']+ s2['other'] + s3['other'] + s4['other'] + s5['other']	 
	 # Assemble return value
	retval = { 'bike' : bike_data, 'ped' : ped_data, 'child' : child_data,
			   'jogger' : jogger_data, 'skater' : skater_data,
			   'wheelchair' : wheelchair_data, 'other' : other_data }
	return retval	
# end_def: read_count_tabs

# Test uber-driver routine:
def test_driver():
	initialize(input_xlsx_fn)
	overview_data = read_overview_tab()
	count_data = read_count_sheets()
	# Here: Have all info needed to assemble and run SQL INSERT query
# end_def: test_driver