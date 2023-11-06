# GUI prototype using Tk 
import tkinter as tk
from tkinter import filedialog as fd

dir_name = ''
pwd = ''

def browse_button():
	global dir_name
	# Allow user to select a directory and store it in global var
	# called folder_path
	dir_name = fd.askdirectory()
	print('dirname = ' + dir_name)
	dirText = tk.Label(master, text=dr_name)
	dirText.grid(row=1, column=2)
#

def process_spreadsheets():
	global dir_name, pwd
	print('dir_name = ' + dir_name)
	print('password = ' + pwd)
#

def do_quit():
	print("That's all folks!")
#

# master = tk.Tk()

# browseButton = tk.Button(master, text="Browse for folder", command=browse_button)
# browseButton.grid(row=1, column=1)

# pwdPromptLabel = tk.Label(master, text="Database password:")
# pwdPromptLabel.grid(row=2, column=1)

# pwdEntry = tk.Entry(master)
# pwdEntry.grid(row=2, column=2)

# runButton = tk.Button("Run", command=process_spreadsheets)
# runButton.grid(row=3, column=1)

# quitButton = tk.Button("Quit", command=do_quit)
# quitButton.grid(row=3, column=2)

# master.mainloop()
# tk.mainloop()

##################################################################################
# import tkinter as tk

def browse_button():
	global dir_name
	# Allow user to select a directory and store it in global var
	# called folder_path
	temp = fd.askdirectory()
	dir_name.set(temp)
	print('dirname = ' + temp)
	# dirLabel(text=temp)
# 

def process_spreadsheets():
	print("Selected folder: " + 'TBD')
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