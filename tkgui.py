# Tk-based GUI for bike-ped traffic count input app.
# Author: Ben Krepp (bkrepp@ctps.org)
 
import tkinter as tk
from tkinter import filedialog as fd

dir_text = None

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
	pwd = pwdEntry.get()
	if pwd == None or pwd == '':
		error_text += 'No password supplied. '
	#
	if dir_text == None or dir_text == '':
		error_text += 'No input folder specified.'
	#
	if error_text != '':
		print(error_text)
		return
	else:
		print("Selected folder: " + dir_text)
		print("DB paassword: %s\n" % pwd)
		# HERE: Call processing driver routine
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
					text="Database password").grid(row=2)
pwdEntry = tk.Entry(master)
pwdEntry.grid(row=2, column=1)

tk.Button(master, 
		  text='Run', command=process_spreadsheets).grid(row=3, 
														 column=0, 
														 sticky=tk.W+tk.E+tk.N+tk.S, 
														 pady=4)
tk.Button(master, 
		  text='Quit', 
		  command=master.quit).grid(row=3, 
									column=1, 
									sticky=tk.W+tk.E+tk.N+tk.S, 
									pady=4)

tk.mainloop()