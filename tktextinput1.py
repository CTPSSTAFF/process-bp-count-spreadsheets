import tkinter as tk
from tkinter import simpledialog

def show_password_dialog():
	global password
	pwd = simpledialog.askstring("Input", "Database password:")
	if pwd:
		password = pwd

parent = tk.Tk()

get_pwd_button = tk.Button(parent, text="Get Password", command=show_password_dialog)
get_pwd_button.pack(padx=20, pady=0)

pwd_label = tk.Label(parent, text="", padx=20, pady=20)
pwd_label.pack()

tk.mainloop()