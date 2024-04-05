import os
import win32com.client as win32

#######################################################################################
##
## In __main__ at the end, change the target folder (in relation to where the scipt is).
## All .xls files in that folder will have an .xlsx created if it does not already exist.
##
#######################################################################################
#	
#	Changes
#	-------
#	21Oct2021	David Ross
#				Adding rough work around for error found with launching Excel and
#				some corruption of the cached python crap.
#
#	25Oct2021	David Ross
#				Making this callable from another python script.

def get_all_files(folder=""):
	path = os.path.dirname(os.path.realpath(__file__))
	path = os.path.join(path, folder)
	os.chdir(path)
	files = []
	#for file in os.listdir(path):	
	for file in sorted(os.listdir(path), key=os.path.getmtime):
		if file.endswith(".xls") and 'template' not in file.lower():
			name, _ = os.path.splitext(os.path.basename(file))
			
			#only add if an equivalent .xlsx does not exist
			if not os.path.exists(os.path.join(path, name + ".xlsx")):
				files.append(os.path.join(path,file))
				print('valid file found:', file, flush=True)
			else:
				print("{} already exists".format(name))

	return files
	
def launchExcel():
	#print("launchExcel")
	try:
		excel = win32.gencache.EnsureDispatch('Excel.Application')
	except AttributeError:
		# Corner case dependencies.
#		import os
		import re
		import sys
		import shutil
		# Remove cache and try again.
		MODULE_LIST = [m.__name__ for m in sys.modules.values()]
		for module in MODULE_LIST:
			if re.match(r'win32com\.gen_py\..+', module):
				del sys.modules[module]
		shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
#		from win32com import client
		excel = win32.gencache.EnsureDispatch('Excel.Application')
	return excel

def convert(fname, excel):	
	#fname = "full+path+to+xls_file"
	print("converting {} to an xlsx".format(os.path.basename(fname)), flush=True )
	wb = excel.Workbooks.Open(fname)
	wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
	wb.Close()                               #FileFormat = 56 is for .xls extension


def xls_2_xlsx(path):
	print("xls_2_xlsx", flush=True)
	files = get_all_files(path)
	excel = launchExcel()
	for file in files:
		convert(file, excel)
	excel.Application.Quit()
		
if __name__ == "__main__":	
	path = input("Tell me the path of the results files:\n> ")
	xls_2_xlxs(path)
	