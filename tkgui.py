# Tk-based GUI for bike-ped traffic count input app.
# Author: Ben Krepp (bkrepp@ctps.org)
 
import tkinter as tk
from tkinter import filedialog as fd
from process_bp_counts import db_initialize, db_shutdown, process_folder

debug = True

dir_text = None
table_text = None

def browse_button():
	global dir_name, dir_text
	# Allow user to select a directory and store it in global var
	# called folder_path
	dir_text = fd.askdirectory()
	dir_name.set(dir_text)
	print('dir_text = ' + dir_text)
#

def process_spreadsheets():
	global dir_text, pwdEntry
	error_text = ''
	db_pwd = pwdEntry.get()

	if db_pwd == None or db_pwd == '':
		error_text += 'No password supplied. '
	#
	if dir_text == None or dir_text == '':
		error_text += 'No input folder specified.'
	#
	if error_text != '':
		print(error_text)
		return
	else:
		# Fill pwdEntry GUI text box with *'s during processing
		pwd_len = len(db_pwd)
		fill = '*'*len(db_pwd)
		pwdEntry.delete(0, pwd_len)
		pwdEntry.insert(0, fill)
		if debug:
			print("Selected folder: " + dir_text)
			print("DB paassword: %s\n" % db_pwd)
		# end_if
		# 
		# Open DB connection
		db_conn = db_initialize(db_pwd)
		if db_conn != None:
			# Call routine to process all XLSXs in the specified folder
			process_folder(dir_text, db_conn)
			print("Returned from call to 'process_folder'.")
			db_shutdown(db_conn)
		else:
			print("Failed to establish connection to database.")
		# end_if
		return
	# end_if
#


master = tk.Tk()
master.title("Load bike-ped count spreadsheets")
# Launch GUI: dimension = 500x100, offset from ULH = 200,200
master.geometry("500x100+200+200")
browseButton = tk.Button(master, 
						 text="Browse for folder", 
						 command=browse_button).grid(row=1, column=0)

dir_name = tk.StringVar()
dirLabel = tk.Label(master,textvariable=dir_name).grid(row=1, column=1)

pwdLabel = tk.Label(master, 
					text="Database passwor:").grid(row=2)
pwdEntry = tk.Entry(master)
pwdEntry.grid(row=2, column=1)

tblLabel = tk.Label(master,
					text="Database table:").grid(row=3)
tblEntry = tk.Entry(master)
tblEntry.grid(row=3, column=1)
					

tk.Button(master, 
		  text='Run', command=process_spreadsheets).grid(row=4, 
														 column=0, 
														 sticky=tk.W+tk.E+tk.N+tk.S, 
														 pady=4)
tk.Button(master, 
		  text='Quit', 
		  command=master.quit).grid(row=4, 
									column=1, 
									sticky=tk.W+tk.E+tk.N+tk.S, 
									pady=4)

tk.mainloop()