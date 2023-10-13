# Prototype #1 of utility to parse spreadsheet of B-P count data,
# and write it to a table in a PostgreSQL database.
#
# Author: Ben Krepp (bkrepp@ctps.org)

import openpyxl
import psycopg

debug = True

# pseudo-constants for coordinates of various fields in 'Overview' sheet
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
comments_coords = 'A12'

input_fn = './xlsx/sample-spreadsheet1.xlsx'

wb = None
overview_sheet = None
count_sheet1 = None
count_sheet2 = None
count_sheet3 = None
count_sheet4 = None
count_sheet5 = None

def initialize(input_fn):
	global wb, overview_sheet, count_sheet1, count_sheet2, count_sheet3, count_sheet4, count_sheet5
	wb = openpyxl.load_workbook(filename = input_fn)
	overview_sheet = wb['Overview']
	count_sheet1 = wb['630-845 AM']
	count_sheet2 = wb['900-1145 AM']
	count_sheet3 = wb['1200-245 PM']
	count_sheet4 = wb['300-545 PM']
	count_sheet5 = wb['600-845 PM']
# end_def


def test_reading_overview():
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
	
	from_st = overview_sheetfrom_st_coords].value
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
		print('temperature = ' + temp)
	
	sky = overview_sheet[sky_coords].value
	if debug:
		print('sky = ' + sky_coords)
	
	comments = overview_sheet(comments_coords).value
	if debug:
		print('comments = '	 + comments)
# end_def

