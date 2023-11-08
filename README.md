# process-bp-count-spreadsheets
Tool to parse Bike-Ped count spreadsheet and write the data it contains to the 'bp_counts' table in the bike-ped 'staging' database.

## Software Dependencies
* Python __openpyxl__ package
* Python __psycopg2__ package
* Python __datetime__ package \(part of Python standard library\)
* Python __glob__ package \(part of Python standard library\)

## Graphical User Interface
A simple graphical user interface \(GUI\) has been developed with 
the __tkinter__ library \(also part of Python standard library\.)  
In addition, some work was done prototyping a GUI with the wxPython
package, but this has been put aside for the time being.
