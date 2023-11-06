# GUI prototype using Tk 
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
	print("Selected folder: " + dir_text)
	print("DB paassword: %s\n" % (pwdEntry.get()))
#


master = tk.Tk()
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
														 sticky=tk.W, 
														 pady=4)
tk.Button(master, 
		  text='Quit', 
		  command=master.quit).grid(row=3, 
									column=1, 
									sticky=tk.W, 
									pady=4)


tk.mainloop()