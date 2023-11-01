#!/usr/bin/env python
"""
Implement a directory (folder) selection dialog.
"""

import wx

class HelloFrame(wx.Frame):
	"""
	A Frame that says Hello World
	"""

	def __init__(self, *args, **kw):
		# ensure the parent's __init__ is called
		super(HelloFrame, self).__init__(*args, **kw)

		# create a panel in the frame
		pnl = wx.Panel(self)
		
		# put some text with a larger bold font on it
		st = wx.StaticText(pnl, label='Select File->Process to select folder to process.')
		font = st.GetFont()
		font.PointSize += 5
		font = font.Bold()
		st.SetFont(font)

		# and create a sizer to manage the layout of child widgets
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(st, wx.SizerFlags().Border(wx.TOP|wx.LEFT, 25))
		pnl.SetSizer(sizer)

		# create a menu bar
		self.makeMenuBar()

		# and a status bar
		self.CreateStatusBar()
		self.SetStatusText("Welcome to bike-ped count processor.")
		



	def makeMenuBar(self):
		"""
		A menu bar is composed of menus, which are composed of menu items.
		This method builds a set of menus and binds handlers to be called
		when the menu item is selected.
		"""

		# Make a file menu with Process and Exit items
		fileMenu = wx.Menu()
		# The "\t..." syntax defines an accelerator key that also triggers
		# the same event
		processItem = fileMenu.Append(-1, "&Process...\tCtrl-P",
				"Help string shown in status bar for this menu item")
		fileMenu.AppendSeparator()
		# When using a stock ID we don't need to specify the menu item's
		# label
		exitItem = fileMenu.Append(wx.ID_EXIT)

		# Now a help menu for the about item
		helpMenu = wx.Menu()
		aboutItem = helpMenu.Append(wx.ID_ABOUT)

		# Make the menu bar and add the two menus to it. The '&' defines
		# that the next letter is the "mnemonic" for the menu item. On the
		# platforms that support it those letters are underlined and can be
		# triggered from the keyboard.
		menuBar = wx.MenuBar()
		menuBar.Append(fileMenu, "&File")
		menuBar.Append(helpMenu, "&Help")

		# Give the menu bar to the frame
		self.SetMenuBar(menuBar)

		# Finally, associate a handler function with the EVT_MENU event for
		# each of the menu items. That means that when that menu item is
		# activated then the associated handler function will be called.
		self.Bind(wx.EVT_MENU, self.OnProcess, processItem)
		self.Bind(wx.EVT_MENU, self.OnExit,	 exitItem)
		self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)


	def OnExit(self, event):
		"""Close the frame, terminating the application."""
		self.Close(True)


	def OnProcess(self, event):
		"""Say hello to the user."""
		# Get selected folder from directory selection dialog
		dialog = wx.DirDialog(self, 'Select folder', '', wx.DD_DEFAULT_STYLE)
		path = ''
		if dialog.ShowModal() == wx.ID_OK:
			path = dialog.GetPath()
			wx.MessageBox('Processing spreadsheets in:\n' + path)
			# Here: Call processing logic
		else:
			wx.MessageBox("No folder selected.")
		#
		

	def OnAbout(self, event):
		"""Display an About Dialog"""
		wx.MessageBox("Select folder containing spreadsheets with bike-ped counts to load.", "",
					  wx.OK|wx.ICON_INFORMATION)


if __name__ == '__main__':
	# When this module is run (not imported) then create the app, the
	# frame, show it, and start the event loop.
	app = wx.App()
	frm = HelloFrame(None, title='Select folder containing spreadsheets to process.')
	frm.Show()
	app.MainLoop()