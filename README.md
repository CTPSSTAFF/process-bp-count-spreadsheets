# process-bp-count-spreadsheets
Tool to parse Bike-Ped count spreadsheet and write the data it contains to the 'bp_counts' table in the bike-ped 'staging' database,
a PostgreSQL database.

## Software Dependencies
* Python __openpyxl__ package
* Python __psycopg2__ package
* Python __datetime__ package \(part of Python standard library\)
* Python __glob__ package \(part of Python standard library\)
* Python __tkinter__ package \(part of Python standard library\)
* Pythong __wxPython__ package

## Programmatic Interface
The module __process\_bp\_counts__ encapsulates the 'processing' logic of this tool.
It has the following logical entrypoints:
* __db\_initialize__ - establish connection with backing database
* __db\_shutdown__ - close connection with backing database
* __process\_folder__ - process all XLSX spreadsheets in a specified folder

## Graphical User Interface
This tool currently supports two graphical user interfaces \(GUIs\):
1. one built using the __tkinter__ package
2. on built using the __wxPython__ package

GUI \#1 is a bit minimalist, and can be run on installations on which only the Python standard library is available.  
GUI \#2 is intended to look 'platform native', but requires the __wxPython__ library to have been installed.  
Both GUI \#1 and GUI \#2 call the programmatic API exposed by __process\_bp\_counts__.
