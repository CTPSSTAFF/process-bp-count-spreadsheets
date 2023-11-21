# wxPython-based GUI for bike-ped traffic count input app.
# Author: Ben Krepp (bkrepp@ctps.org)

import wx, wx.html, sys
from process_bp_counts import db_initialize, db_shutdown, process_folder

# Code for the application's GUI begins here.
#
aboutText = """<p>Help text for this program is TBD.<br>
This program is running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

class HtmlWindow(wx.html.HtmlWindow):
	def __init__(self, parent, id, size=(600,400)):
		wx.html.HtmlWindow.__init__(self,parent, id, size=size)
		if "gtk2" in wx.PlatformInfo:
			self.SetStandardFonts()
	# end_def __init__()

	def OnLinkClicked(self, link):
		wx.LaunchDefaultBrowser(link.GetHref())
	# end_def OnLinkClicked()
# end_class HtmlWindow

class AboutBox(wx.Dialog):
	def __init__(self):
		wx.Dialog.__init__(self, None, -1, "About the bike-ped traffic count input tool.",
						   style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.TAB_TRAVERSAL)
		hwin = HtmlWindow(self, -1, size=(400,200))
		vers = {}
		vers["python"] = sys.version.split()[0]
		vers["wxpy"] = wx.VERSION_STRING
		hwin.SetPage(aboutText % vers)
		btn = hwin.FindWindowById(wx.ID_OK)
		irep = hwin.GetInternalRepresentation()
		hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
		self.SetClientSize(hwin.GetSize())
		self.CentreOnParent(wx.BOTH)
		self.SetFocus()
	# end_def __init__()
# end_class AboutBox

# This is the class for the main GUI itself.
class Frame(wx.Frame):
	# Name of directory containing XLSX files to be read
	inputDirName = ''
	# DB table name
	dbTableName = ''
	# DB password
	db_pwd = '' 
	
	def __init__(self, title):
		wx.Frame.__init__(self, None, title=title, pos=(250,250), size=(600,320),
						  style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)
		self.Bind(wx.EVT_CLOSE, self.OnClose)

		menuBar = wx.MenuBar()
		menu = wx.Menu()
		m_exit = menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
		self.Bind(wx.EVT_MENU, self.OnClose, m_exit)
		menuBar.Append(menu, "&File")
		menu = wx.Menu()
		m_about = menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
		self.Bind(wx.EVT_MENU, self.OnAbout, m_about)
		menuBar.Append(menu, "&Help")
		self.SetMenuBar(menuBar)
		
		self.statusbar = self.CreateStatusBar()

		panel = wx.Panel(self)
		box = wx.BoxSizer(wx.VERTICAL)
		box.AddSpacer(20)
			  
		m_select_input_dir = wx.Button(panel, wx.ID_ANY, "Select input folder")
		m_select_input_dir.Bind(wx.EVT_BUTTON, self.OnSelectInputDir)
		box.Add(m_select_input_dir, 0, wx.CENTER)
		box.AddSpacer(20)		
		
		# Placeholder for name of selected input folder; it is populated in OnSelectInputDir(). 
		self.m_dirText = wx.StaticText(panel, -1, " ")
		self.m_dirText.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
		self.m_dirText.SetSize(self.m_dirText.GetBestSize())
		box.Add(self.m_dirText, 0, wx.ALL, 10)	 
		box.AddSpacer(20)
		
		l1 = wx.StaticText(panel, -1, "Database table name:") 
		box.Add(l1)
		box.AddSpacer(20)
		
		# Name of destination database table
		self.m_tblText = wx.TextCtrl(panel, value="ctps_bp_counts_staging") 
		self.m_tblText.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
		self.m_tblText.SetSize(self.m_tblText.GetBestSize())
		box.Add(self.m_tblText, 0, wx.CENTER)
		box.AddSpacer(20)
		
		# Button to popup wx.TextEntryDialog to collect DB pwd
		m_db_pwd = wx.Button(panel, wx.ID_ANY, "Enter database password")
		m_db_pwd.Bind(wx.EVT_BUTTON, self.OnGetDbPwd)
		box.Add(m_db_pwd, 0, wx.CENTER)
		box.AddSpacer(20)
		
		m_run = wx.Button(panel, wx.ID_ANY, "Load bike/ped counts")
		m_run.Bind(wx.EVT_BUTTON, self.OnRun)
		box.Add(m_run, 0, wx.CENTER)		
		
		panel.SetSizer(box)
		panel.Layout()
	# end_def __init__()
		
	def OnClose(self, event):
		dlg = wx.MessageDialog(self, 
			"Do you really want to close this application?",
			"Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
		result = dlg.ShowModal()
		dlg.Destroy()
		if result == wx.ID_OK:
			self.Destroy()
	# end_def OnClose()

	def OnGetDbPwd(self, event):
		frame = wx.Frame(None, -1, 'win.py')
		frame.SetSize(0,0,200,50)
		dlg = wx.TextEntryDialog(None, 'Please enter database password below:',
			 'Database is password protected', '',
			 style = wx.TextEntryDialogStyle|wx.TE_PASSWORD)
		dlg.ShowModal()
		self.db_pwd = str(dlg.GetValue())
		dlg.Destroy()
		frame.Destroy()
	# end_def OnGetDbPwd()	 
	
	def OnSelectInputDir(self, event):
		frame = wx.Frame(None, -1, 'win.py')
		frame.SetSize(0,0,200,50)
		dlg = wx.DirDialog(None, "Select input folder", "",
						   wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
		dlg.ShowModal()
		self.inputDirName = dlg.GetPath()
		self.m_dirText.SetLabel("Selected input folder: " + self.inputDirName)
		dlg.Destroy()
		frame.Destroy()
	# end_def OnSelectInputDir()	
	
	def OnRun(self, event):
		msg = ''
		if self.inputDirName == '':
			msg += 'Input folder name not specified. '
		#
		if self.db_pwd == '':
			msg += 'Database password not supplied.'
		#
		if len(msg) > 0:
			msg = 'Error(s): ' + msg
			dlg = wx.MessageDialog(self, msg, 'Error(s)', wx.OK)
			result = dlg.ShowModal()
			dlg.Destroy()
			return
		#
		# Here: We have folder name and DB password.
		dlg = wx.MessageDialog(self, 
			"Confirm that you want to run the bike-ped count loading tool.",
			"Confirm: OK/Cancel", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
		result = dlg.ShowModal()
		dlg.Destroy()
		if result == wx.ID_OK:
			# 1. Establish database connection
			db_conn = db_initialize(self.db_pwd)
			if db_conn != None:
				# 2. Call routine to process XLSX files in specified folder
				# *** TBD: DB table name parm (parm #2) is hard-wired for the moment.
				process_folder(self.inputDirName, 'ctps_bp_counts_staging', db_conn)
				print("Returned from call to 'process_folder'.")
				# 3. Close database connection
				db_shutdown(db_conn)
				message = "Bicycle/Pedestrian count data loaded."
			else: 
				message = "Failed to establish database connection."
			# end_if
			caption = "Bicycle/Pedestrian Traffic Count Loader"
			dlg = wx.MessageDialog(None, message, caption, wx.OK | wx.ICON_INFORMATION)
			dlg.ShowModal()
			dlg.Destroy()
		else:
			pass
			# self.Destroy()
		# end_if
	# end_def OnRun()

	def OnAbout(self, event):
		dlg = AboutBox()
		dlg.ShowModal()
		dlg.Destroy() 
	# end_def OnAbout()
# end_class Frame

# The code for the GUI'd application itself begins here.
#
app = wx.App(redirect=True)	  # Error messages go to popup window
top = Frame("Load Bicycle/Pedestrian Traffic Counts")
top.Show()
app.MainLoop()
