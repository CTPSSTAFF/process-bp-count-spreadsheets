import tkinter as tk
from tkinter import *
from tkinter import filedialog as fd


def browse_button():
	global dir_name, label1
	# Allow user to select a directory and store it in global var
	# called folder_path
	temp = fd.askdirectory()
	dir_name.set(temp)
	print('dirname = ' + temp)
	label1.text = temp
#

root = Tk()

button1 = tk.Button(text="Browse for folder", command=browse_button)
button1.pack(fill=tk.X)

dir_name = tk.StringVar()
label1 = tk.Label(master=root,textvariable=dir_name)
label1.pack(fill=tk.X)

button2 = tk.Button(text="Enter database password:")
button2.pack(fill=tk.X)

mainloop()