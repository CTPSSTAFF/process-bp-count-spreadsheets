# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
from openpyxl.formula import Tokenizer
import psycopg
import datetime

debug = True

# input_xlsx_fn = './xlsx/sample-spreadsheet1.xlsx'
input_xlsx_fn = './xlsx/template-spreadsheet.xlsx'

# Pseudo-constants for coordinates of various fields in 'Overview' sheet.
date_coords = 'C2'
temperature_coords = 'C3'
sky_coords = 'C4'
#
bp_loc_id_coords = 'R3'
count_id_coords = 'R4'

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

# List of range or row numbers with data in each of the count sheets
data_sheet_rows = range(2, 14) # i.e., 2 to 13

# Lists of 'key values' for rows in each of the data sheets,
# i.e., the name of the relevant column in the datbase counts table
keys_sheet_1 = [ 'cnt_0600', 'cnt_0615', 'cnt_0630', 'cnt_0645',
				 'cnt_0700', 'cnt_0715', 'cnt_0730', 'cnt_0745',
				 'cnt_0800', 'cnt_0815', 'cnt_0830', 'cnt_0845' ]
				 
keys_sheet_2 = [ 'cnt_0900', 'cnt_0915', 'cnt_0930', 'cnt_0945',
				 'cnt_1000', 'cnt_1015', 'cnt_1030', 'cnt_1045',
				 'cnt_1100', 'cnt_1115', 'cnt_1130', 'cnt_1145' ]

keys_sheet_3 = [ 'cnt_1200', 'cnt_1215', 'cnt_1230', 'cnt_1245',
				 'cnt_1300', 'cnt_1315', 'cnt_1330', 'cnt_1345',
				 'cnt_1400', 'cnt_1415', 'cnt_1430', 'cnt_1445' ]

keys_sheet_4 = [ 'cnt_1500', 'cnt_1515', 'cnt_1530', 'cnt_1545',
				 'cnt_1600', 'cnt_1615', 'cnt_1630', 'cnt_1645',
				 'cnt_1700', 'cnt_1715', 'cnt_1730', 'cnt_1745' ]
				 
keys_sheet_5 = [ 'cnt_1800', 'cnt_1815', 'cnt_1830', 'cnt_1845',
				 'cnt_1900', 'cnt_1915', 'cnt_1930', 'cnt_1945',
				 'cnt_2000', 'cnt_2015', 'cnt_2030', 'cnt_2045' ]

keys_sheet_6 = [ 'cnt_2100', 'cnt_2115', 'cnt_2130', 'cnt_2145',
				 'cnt_2200', 'cnt_2215', 'cnt_2230', 'cnt_2245',
				 'cnt_2300', 'cnt_2315', 'cnt_2330', 'cnt_2345' ]

keys_sheet_7 = [ 'cnt_0000', 'cnt_0015', 'cnt_0030', 'cnt_0045',
				 'cnt_0100', 'cnt_0115', 'cnt_0130', 'cnt_0145',
				 'cnt_0200', 'cnt_0215', 'cnt_0230', 'cnt_0245' ]

keys_sheet_8 = [ 'cnt_0300', 'cnt_0315', 'cnt_0330', 'cnt_0345',
				 'cnt_0400', 'cnt_0415', 'cnt_0430', 'cnt_0445',
				 'cnt_0500', 'cnt_0515', 'cnt_0530', 'cnt_0545' ]



wb = None
overview_sheet = None
count_sheet_1 = None	# 6:00-8:45 AM
count_sheet_2 = None	# 9:00-11:45 AM
count_sheet_3 = None	# 12:00-2:45 PM
count_sheet_4 = None	# 3:00-5:45 PM
count_sheet_5 = None	# 6:00-8:45 PM
count_sheet_6 = None	# 9:00-11:45 PM
count_sheet_7 = None	# 12:00-2:45 AM
count_sheet_8 = None	# 3:00-5:45 AM
columns_sheet = None	# Sheet containing lookup tables (no longer used)


def initialize(input_fn):
	global wb, overview_sheet, count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4
	global count_sheet_5, count_sheet_6, count_sheet_7, count_sheet_8
	
	wb = openpyxl.load_workbook(filename = input_fn)
	overview_sheet = wb['Overview']
	count_sheet_1 = wb['600-845 AM']
	count_sheet_2 = wb['900-1145 AM']
	count_sheet_3 = wb['1200-245 PM']
	count_sheet_4 = wb['300-545 PM']
	count_sheet_5 = wb['600-845 PM']
	count_sheet_6 = wb['900-1145 PM']
	count_sheet_7 = wb['1200-245 AM']
	count_sheet_8 = wb['300-545 AM']
	columns_sheet = wb['Columns']
# end_def


# read_overview_sheet: read and parse data from 'Overview' sheet
#
def read_overview_sheet():
	global overview_sheet, debug
	
	bp_loc_id = overview_sheet[bp_loc_id_coords].value
	count_id = overview_sheet[bp_count_id_coords].value
	
	loc_desc = overview_sheet[loc_desc_coords].value
	if loc_desc == None:
		loc_desc = ''
	elif loc_desc == 'Other':
		loc_desc = overview_sheet[loc_desc_other_coords].value
	# end_if
	
	
	date_raw = overview_sheet[date_coords].value
	# As best we can tell, the 'value' is in yyyy-mm-dd hh:mm:ss format.
	# Extract just the 'date' part.
	if date_raw == None:
		# Temp workaround, for now
		date_raw = '10/23/2023'
		date_cooked = date_raw
	else:
		date_cooked = datetime.datetime.strftime(date_raw, '%m-%d-%Y')
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
		print('bp_loc_id = ' + bp_loc_id)
		print('count_id = ' + count_id)
		print('date	 = ' + date_cooked)
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
		
	# Assemble return value: dict of information harvested from overview table
	retval = { 'bp_loc_id' : bp_loc_id, 'count_id' : count_id, 'date' : date_cooked, 
			   'street_1' : street_1, 'street_1_dir' : street_1_dir,
			   'street_2' : street_2, 'street_2_dir' : street_2_dir,   
			   'temperature' : temperature, 'sky' : sky,
			   'comments' : comments }
	return retval
# end_def: read_overview_sheet


# read_count_sheet: Read data from _one_ count sheet.
#
# Parameters:
#	  count_sheet - workbook count sheet to be read
#	  rows - range of rows to be read in count sheet
#
# Return value:
#	A 3-level data structure.
#	The top-level has the keys: 'bike', 'ped', 'child', 'jogger',
#								'skater', 'wheelchair', and 'other
#  The second level: each of the topl-level dicts contains a list
#  each of whose 96 elements are dicts.
#  Each of these dicts has two keys: 'k' and 'v'.
#  The value of the 'k' key is the name of a database 'count' column,
#  e.g., 'cnt_1030'; the value of the 'v' key is the corresponding
#  count, which may be None, i.e., NULL.
#
def read_count_sheet(count_sheet, rows, row_keys):
	global debug
	
	bike_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = bike_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		bike_temp.append(dtmp)
	#
	
	ped_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = ped_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		ped_temp.append(dtmp)
	#
	
	child_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = child_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		child_temp.append(dtmp)
	#
	
	jogger_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = jogger_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		jogger_temp.append(dtmp)
	#
	
	skater_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = skater_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		skater_temp.append(dtmp)
	#
	
	wheelchair_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = wheelchair_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		wheelchair_temp.append(dtmp)
	#
	
	other_temp = []
	for (row_ix, row_key) in zip(rows, row_keys):
		ix = other_col + str(row_ix)
		tmp = count_sheet[ix].value
		try:
			val = int(tmp)
		except:
			val = None
		#
		dtmp = { 'k' : row_key, 'v' : val }
		other_temp.append(dtmp)
	#
	
	if debug:
		print('Bike counts:')
		for c in bike_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Ped counts:')
		for c in ped_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Child counts:')
		for c in child_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Jogger counts:')
		for c in jogger_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Skater counts:')
		for c in skater_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Wheelchair counts:')
		for c in wheelchair_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
		print('Other counts:')
		for c in other_temp:
			s = c['k'] + ' : ' 
			s += str(c['v']) if c['v'] != None else 'NULL'
			print(s)
		#
	# end_if debug
	
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
	global count_sheet_1, count_sheet_2, count_sheet_3, count_sheet_4
	global count_sheet_5, count_sheet_6, count_sheet_7, count_sheet_8
	global keys_sheet_1, keys_sheet_2, keys_sheet_3, keys_sheet_4
	global keys_sheet_5, keys_sheet_6, keys_sheet_7, keys_sheet_8
	global data_sheet_rows
	s1 = read_count_sheet(count_sheet_1, data_sheet_rows, keys_sheet_1)
	s2 = read_count_sheet(count_sheet_2, data_sheet_rows, keys_sheet_2)
	s3 = read_count_sheet(count_sheet_3, data_sheet_rows, keys_sheet_3)
	s4 = read_count_sheet(count_sheet_4, data_sheet_rows, keys_sheet_4)
	s5 = read_count_sheet(count_sheet_5, data_sheet_rows, keys_sheet_5)
	s6 = read_count_sheet(count_sheet_6, data_sheet_rows, keys_sheet_6)
	s7 = read_count_sheet(count_sheet_7, data_sheet_rows, keys_sheet_7)
	s8 = read_count_sheet(count_sheet_8, data_sheet_rows, keys_sheet_8)
	
	# Assemble count data from all sheets.
	# Note that the order of the sheets, according to the 24-hour clock is: 7, 8, 1, 2, 3, 4, 5, 6
	bike_data = s7['bike'] + s8['bike'] + s1['bike'] + s2['bike'] + s3['bike'] + s4['bike'] + s5['bike'] + s6['bike']
	ped_data = s7['ped'] + s8['ped'] + s1['ped']+ s2['ped'] + s3['ped'] + s4['ped'] + s5['ped'] + s6['ped']
	child_data = s7['child'] + s8['child'] + s1['child']+ s2['child'] + s3['child'] + s4['child'] + s5['child'] + s6['child']
	jogger_data = s7['jogger'] + s8['jogger'] + s1['jogger']+ s2['jogger'] + s3['jogger'] + s4['jogger'] + s5['jogger'] + s6['jogger']
	skater_data = s7['skater'] + s8['skater'] +	 s1['skater'] + s2['skater'] + s3['skater'] + s4['skater'] + s5['skater'] + s6['skater']
	wheelchair_data = s7['wheelchair']+ s2['wheelchair'] + s3['wheelchair'] + s4['wheelchair'] + s5['wheelchair'] + s6['wheelchair']
	other_data = s7['skater']+ s8['skater'] + s1['other']+ s2['other'] + s3['other'] + s4['other'] + s5['other'] + s6['other'] 
	# Assemble return value
	retval = { 'bike' : bike_data, 'ped' : ped_data, 'child' : child_data,
			   'jogger' : jogger_data, 'skater' : skater_data,
			   'wheelchair' : wheelchair_data, 'other' : other_data }
	return retval	
# end_def: read_count_sheets

# run_insert_queries: run INSERT QUERIES to insert count data into staging counts table
#
# parameters: overview - data harvested from overview sheet
#			  counts - data harvested from count sheets
#
def run_insert_queries(overview, counts):
	pass
	# *** TO BE WRITTEN
# end_def run_insert_queries

# Test uber-driver routine:
def test_driver(xlsx_fn):
	initialize(xlsx_fn)
	overview_data = read_overview_sheet()
	count_data = read_count_sheets()
	# Here: Have all info needed to assemble and run SQL INSERT INTO query
	run_insert_queries(overview_data, count_data)
# end_def: test_driver

# Test driver for only reading 'Overview' sheet
def test_driver_overview(xlsx_fn):
	initialize(xlsx_fn)
	overview_data = read_overview_sheet()
# end_def

# Test driver for only reading count data
def test_driver_counts(xlsx_fn):
	initialize(xlsx_fn)
	count_data = read_count_sheets()
# end_def: test_driver_counts
