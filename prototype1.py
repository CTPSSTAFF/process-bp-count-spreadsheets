# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
import psycopg2
import datetime

# Debug toggles
debug_read_overview = True
debug_read_counts = False
debug_query_string = True

# input_xlsx_fn = './xlsx/template-spreadsheet.xlsx'
input_xlsx_fn = './xlsx/sample-spreadsheet3.xlsx'

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
from_st_name_coords = 'D7'
from_st_dir_coords = 'D8'
to_st_name_coords = 'D9'
to_st_dir_coords = 'D10'
#
loc_type_coords = 'D11'
muni_coords = 'D12'
muni_other_coords = 'F12'
#
description_coords = 'J4'

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

# Pseudo-constant for 'temperature_not_recorded'
temp_not_recorded = -99

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


# bail_out: report fatal error and exit
#
def bail_out(msg):
	print('Fatal error:')
	print('\t' + msg)
	exit()
#


def spreadsheet_initialize(input_fn):
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
# end_def spreadsheet_initialize


# read_overview_sheet: read and parse data from 'Overview' sheet
#
# return value - dict containing the following keys:
#	  'bp_loc_id', 'count_id', 'date', dow_str',
#	  'from_st_name', 'from_st_dir','to_st_name', 'to_st_dir', 
#	  'temperature', 'sky', 'description'
#
def read_overview_sheet():
	global overview_sheet, debug_read_overview
	
	# bp_loc_id: type is INTEGER
	bp_loc_id = overview_sheet[bp_loc_id_coords].value
	if bp_loc_id == None:
		bail_out("count location ID missing")
	#
	bp_loc_id = bp_loc_id
	
	# count_id: type is STRING
	count_id = overview_sheet[count_id_coords].value
	if count_id == None:
		bail_out("count_id missing.")
	#
	count_id = count_id
	
	loc_desc = overview_sheet[loc_desc_coords].value
	if loc_desc == None:
		loc_desc = ''
	elif loc_desc == 'Other':
		loc_desc = overview_sheet[loc_desc_other_coords].value
	# end_if
	
	
	date_raw = overview_sheet[date_coords].value
	if date_raw == None:
		bail_out("Date missing.")
	#
	# As best we can tell, the 'value' in the spreadsheet is in yyyy-mm-dd hh:mm:ss format.
	# Extract just the 'date' part.
	# Convert to PostgreSQL date format: yyyy-mm-dd.:
	date_cooked = datetime.datetime.strftime(date_raw, "%Y-%m-%d")
	
	dow_ix = datetime.date.weekday(date_raw)
	# Python's datetime.date.weekday returns a 0-based index of the day-of-the-week for a given date object.
	# 0 corresponds to 'Monday', 1 to 'Tuesday, etc.
	# The cnt_dow is encoded in the count database as a 3-character string.
	dow_translation_table = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
	dow_str = dow_translation_table[dow_ix]
	
	loc_type = overview_sheet[loc_type_coords].value
	
	muni = overview_sheet[muni_coords].value
	if muni == 'Other':
		muni = overview_sheet[muni_other_coords].value
	#	 
	
	fac_name = overview_sheet[fac_name_coords].value
	if fac_name == 'Other':
		fac_name = overview_sheet[fac_name_other_coords].value
	#
	
	from_st_name = overview_sheet[from_st_name_coords].value
	if from_st_name == None:
		from_st_name = ''
	from_st_dir = overview_sheet[from_st_dir_coords].value
	to_st_name = overview_sheet[to_st_name_coords].value
	if to_st_name == None:
		to_st_name = ''
	to_st_dir = overview_sheet[to_st_dir_coords].value

	# temperature: type is INTEGER
	temperature = overview_sheet[temperature_coords].value
	if temperature == None:
		temperature = temp_not_recorded 
	else:
		temperature = temperature
	#
	
	# sky: type is INTEGER - string converted to integer
	sky = overview_sheet[sky_coords].value
	if sky == 'Sunny':
		sky = '1'
	elif sky == 'Partly Cloudy':
		sky = '2'
	elif sky == 'Overcast':
		sky = '3'
	elif sky == 'Precipitation':
		sky = '4'
	elif sky == 'No Data':
		sky = '99'
	elif sky == None:
		sky = '99'
	else:
		sky = '99'
	# end_if
	
	description = overview_sheet[description_coords].value
	if description == None:
		description = ''
	#
	
	if debug_read_overview:
		print('bp_loc_id = ' + str(bp_loc_id))
		print('count_id = ' + count_id)
		print('date	 = ' + date_cooked)
		print('dow = ' + dow_str)
		print('location description = ' + loc_desc)
		print('location type = ' + loc_type)
		print('municipality = ' + muni)
		print('facility name = ' + fac_name)
		print('from street name = ' + from_st_name)
		print('from street direction = ' + from_st_dir)
		print('to street name = ' + to_st_name)
		print('to street direction = ' + to_st_dir)
		print('temperature = ' + str(temperature))
		print('sky = ' + str(sky))
		print('description = '	 + description)
	# end_if 
		
	# Assemble return value: dict of information harvested from overview table
	retval = { 'bp_loc_id' : bp_loc_id, 'count_id' : count_id, 'date' : date_cooked, 'dow' : dow_str,
			   'from_st_name' : from_st_name, 'from_st_dir' : from_st_dir,
			   'to_st_name' : to_st_name, 'to_st_dir' : to_st_dir, 
			   'temperature' : temperature, 'sky' : sky,
			   'description' : description }
	return retval
# end_def: read_overview_sheet


# read_count_sheet: Read data from _one_ count sheet.
#
# Parameters:
#	  count_sheet - workbook count sheet to be read
#	  rows - range of rows to be read in count sheet
#			 At one point, a different number of rows was to be read from the first sheets
#			 than the others, but this has now been changed. Retaining this parameter in
#			 case things change again.
#
# Return value:
#	A 3-level data structure.
#	The top-level has the keys: 'bike', 'ped', 'child', 'jogger',
#								'skater', 'wheelchair', and 'other'
#  The second level: each of the topl-level dicts contains a list
#  12 elements are dicts - one for each of the 15 minute chunks of time
#  in the time period covered by the sheet.
#
#  Each of these dicts has two keys: 'k' and 'v'.
#  The value of the 'k' key is the name of a database 'count' column,
#  e.g., 'cnt_1030'; the value of the 'v' key is the corresponding
#  count, which may be None, i.e., NULL.
#
def read_count_sheet(count_sheet, rows, row_keys):
	global debug_read_counts
	
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
	
	if debug_read_counts:
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


# read_count_sheets - read data from all count sheets.
# This is a 'driver routine' that calls read_count_sheet for 
# each of the 8 individual 'count' sheets' and assembles
# a single return value for all count data.
#
# Return value:
#	A 3-level data structure.
#	The top-level has the keys: 'bike', 'ped', 'child', 'jogger',
#								'skater', 'wheelchair', and 'other'
#  The second level: each of the topl-level dicts contains a list
#  each of whose 96 elements are dicts - one for each of the
#  15 minute chunks of time in a 24-hour day.
#
#  Each of these dicts has two keys: 'k' and 'v'.
#  The value of the 'k' key is the name of a database 'count' column,
#  e.g., 'cnt_1030'; the value of the 'v' key is the corresponding
#  count, which may be None, i.e., NULL.
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
	other_data = s7['other']+ s8['other'] + s1['other']+ s2['other'] + s3['other'] + s4['other'] + s5['other'] + s6['other'] 
	# Assemble return value
	retval = { 'bike' : bike_data, 'ped' : ped_data, 'child' : child_data,
			   'jogger' : jogger_data, 'skater' : skater_data,
			   'wheelchair' : wheelchair_data, 'other' : other_data }
	return retval	
# end_def: read_count_sheets


# geenrate_insert query: generate text of a single INSERT INTO query for B-P count data into staging counts table;
#					this is called once per 'mode' for any given count_id
#
# parameters: overview - data harvested from overview sheet
#			  count - count data for a single mode harvested from counts sheets
#			  table_name - name of table into which to insert data
#			  mode - mode of travel, e.g., 'bike' or 'ped'
#
# return value: text of SQL INSERT INTO query string
#
def generate_insert_query(overview, count, table_name, mode):
	global debug_query_string
	
	# Common fields, from 'overview' sheet
	bp_loc_id = overview['bp_loc_id']
	count_id = overview['count_id']
	count_date = overview['date']
	count_dow = overview['dow']
	from_st_name = overview['from_st_name']
	from_st_dir = overview['from_st_dir']
	to_st_name = overview['to_st_name']
	to_st_dir = overview['to_st_dir']
	temperature = overview['temperature']
	sky = overview['sky']
	description = overview['description']
	
	# The following is, admittedly, a quick hack
	if mode == 'bike':
		count_type = 'B'
	elif mode == 'ped':
		count_type = 'P'
	elif mode == 'child':
		count_type = 'C'
	elif mode == 'jogger':
		count_type = 'J'
	elif mode == 'skater':
		count_type = 'S'
	elif mode == 'wheelchair':
		count_type = 'W'
	elif mode == 'other':
		count_type = 'O'
	else:
		bail_out("Unrecognized mode (" + mode + ") in 'run_insert_query'.")
	# end_if
	
	# Prep for constructing query string: Get lists of 'overview' keys with non-Null values, and (2) those values
	overview_keys_list = []
	overview_vals_list = []
	
	# Note that bp_loc_id has type INTEGER;
	# count_id, count_date, count_dow, and count_type have type STRING.
	overview_keys_list.append('bp_loc_id')
	overview_keys_list.append('count_id')
	overview_keys_list.append('count_date')
	overview_keys_list.append('count_dow')
	overview_keys_list.append('count_type')
	overview_vals_list.append(str(bp_loc_id))
	overview_vals_list.append("'" + count_id + "'")
	overview_vals_list.append("'" + count_date + "'")
	overview_vals_list.append("'" + count_dow + "'")
	overview_vals_list.append("'" + count_type + "'")
	
	# from_st_name and from_st_dir - these are of type STRING
	if from_st_name != '':
		overview_keys_list.append('from_st_name')
		overview_vals_list.append("'" + from_st_name + "'")
	#
	if from_st_dir != '':
		overview_keys_list.append('from_st_dir')
		overview_vals_list.append("'" + from_st_dir + "'")
	#
	
	# to_st_name and to_st_dir - these are of type STRING
	if to_st_name != '':
		overview_keys_list.append('to_st_name')
		overview_vals_list.append("'" + to_st_name + "'")
	#
	if to_st_dir != '':
		overview_keys_list.append('to_st_dir')
		overview_vals_list.append("'" + to_st_dir  + "'")
	#
	
	# temperature and sky: these are of type INTEGER
	if temperature != temp_not_recorded:
		overview_keys_list.append('temperature')
		overview_vals_list.append(int(temperature))
	#
	if sky != '':
		overview_keys_list.append('sky')
		overview_vals_list.append(sky)
	#
	
	# description
	# Escape any single quotes in description string
	if description != '':
		overview_keys_list.append('description')
		description_cooked = description.replace("'", "''")
		overview_vals_list.append("'" + description_cooked + "'")
	#
	
	# DEBUG
	# if debug_query_string:
	#	for item in overview_vals_list:
	#		print(str(item))
	#	#
	#
	
	overview_keys_string = ', '.join(overview_keys_list)
	overview_vals_string = ', '.join(overview_vals_list)
	
	
	# Prep for constructing query string: Get lists of (1) keys with non-Null values and (2) those values from 'count'.
	# While we're at it, calculate cnt_total
	count_keys_list =[]
	count_vals_list = []
	cnt_total = 0
	for i in count:
		if i['v'] != None:
			count_keys_list.append(i['k'])
			count_vals_list.append(str(i['v']))
			cnt_total += i['v']
		#
	#
	count_keys_list.append('cnt_total')
	count_vals_list.append(str(cnt_total))
	count_keys_string = ', '.join(count_keys_list)
	count_vals_string = ', '.join(count_vals_list)
	
	# Assemble query string
	#
	part1 = "INSERT INTO " + table_name + " ("
	part1 += overview_keys_string
	part1 += ", "
	
	# List of 'count' columns for which we have data for this mode
	part2 =	 " " + count_keys_string + " ) "
	
	part3 = "VALUES ( "
	part3 += overview_vals_string
	
	# List of values for 'count' coulumns for which we have data for this mode
	part4 =	 ", " + count_vals_string + " );"
	query_string = part1 + part2 + part3 + part4
	
	if debug_query_string:
		print('Query string for mode ' + mode + ':')
		print(query_string)
	#
	return query_string
# end_generate_insert_query


def run_insert_query(query_string, db_conn, db_cursor):
	try:
		db_cursor.execute(query_str)
	except:
		db_conn.rollback()
	else:
		db_conn.commit()
	#
# end_def run_insert_query


# run_insert_queries: driver routine for running INSERT QUERIES to insert B-P count data into staging counts table
#
# parameters: overview - data harvested from overview sheet
#			  counts - data harvested from count sheets
#			  table_name - name of table into which to insert data
#			  db_cursor - database cursor
#
def run_insert_queries(overview, counts, table_name, db_conn, db_cursor):
	# For each of the 'modes' (e.g., 'bike', 'ped', etc.) only assemble and execute
	# an INSERT INTO query if there is at least one) real data value for that mode
	# in the input spreadsheet.
	#
	for mode in ['bike', 'ped', 'child', 'jogger', 'skater', 'wheelchair', 'other']:
		c = counts[mode]
		t = [x['v'] == None for x in c]
		if any(y == True for y in t):
			query_string = generate_insert_query(overview, c, table_name, mode)
			run_insert_query(query_string, db_conn, db_cursor)
		#
	 #
# end_def run_insert_queries


def db_initialize(parm, db_pwd):
	# The last two parameters to the 'connect' call, per:
	# https://stackoverflow.com/questions/59190010/psycopg2-operationalerror-fatal-unsupported-frontend-protocol-1234-5679-serve
	#
	if parm == 'office':
		try:
			conn = psycopg2.connect(dbname="mpodata", 
									host="appsrvr3.ad.ctps.org",
									port=5433,
									user="postgres", 
									password=db_pwd,
									sslmode="disable",
									gssencmode="disable")
			retval = conn
		except psycopg2.Error as e:
			print('Error code: ' + e.pgcode)
			print(e.pgerror)
			retval = None
	else:
		try:
			conn = psycopg2.connect(dbname="postgres", 
									host="localhost",
									port=5432,
									user="postgres", 
									password=db_pwd,
									sslmode="disable",
									gssencmode="disable")
			retval = conn
		except psycopg2.Error as e:
			print('Error code: ' + e.pgcode)
			print(e.pgerror)
			retval = None
	# end_if
	return retval
# end_df db_initialize


# Test uber-driver routine:
def test_driver(xlsx_fn, table_name, db_parm, db_pwd):
	spreadsheet_initialize(xlsx_fn)
	overview_data = read_overview_sheet()
	count_data = read_count_sheets()
	# Here: Have read info from the spreadsheet needed to assemble and run SQL INSERT INTO query
	# Initialize for database operations
	db_conn = database_initialize(db_parm, db_pwd)
	if db_conn != None:
		db_cursor = db_conn.cursor()
		run_insert_queries(overview_data, count_data, table_name, db_conn, db_cursor)
		# Shutdown database connection
		db_cursor.close()
		db_conn.close()
	# end_if
# end_def: test_driver


# Test driver for only reading data from 'Overview' sheet
def test_driver_overview(xlsx_fn):
	spreadsheet_initialize(xlsx_fn)
	overview_data = read_overview_sheet()
# end_def

# Test driver for only reading count data from spreadsheet
def test_driver_counts(xlsx_fn):
	spreadsheet_initialize(xlsx_fn)
	count_data = read_count_sheets()
	return count_data
# end_def: test_driver_counts

# Test driver for database connection
def test_driver_db(db_parm, db_pwd):
	db_conn = database_initialize(db_parm, db_pwd)
	if db_conn != None:
		db_cursor = db_conn.cursor()
		db_cursor.close()
		db_conn.close()
	#
# end_def
