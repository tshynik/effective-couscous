## Script to process and consolidate DocuSign exports.

## inspiration from https://sukhbinder.wordpress.com/2014/05/07/simple-and-minimal-file-selection-gui-with-standard-python/
# stuff got renamed in py3: https://stackoverflow.com/questions/28590669/tkinter-tkfiledialog-doesnt-exist

import tkinter as tk
from tkinter import filedialog as tkfile
import csv

## TODO: It'd be nice if I could get a Windows Explorer to open, pointed at the directory where the file was created. This does not want to work, however.
# https://stackoverflow.com/questions/36449361/tkinter-way-to-open-a-directory-window-in-windows-explorer
#from os import system, path
#import subprocess
def openfolder(root, filepath):
	messaging(filepath)
	folderpath = '\\'.join(filepath.split('\\')[0:-1])
	messaging(folderpath)
	os.system("start " + folderpath)

## Specify formats the filepicker allows here.
myFormats = [
	('Comma Separated File','*.csv')#,
	#('Tab-Delimited File','*.txt'),
	#('Tab-Delimited File','*.tsv')
	]

## Print messages to both the command prompt (if running it from there) and the messaging box in the GUI.
def messaging(msg_text):
	#results.configure(state="enabled") #TODO: This doesn't work, maybe wrong keyword?
	results.insert("end", msg_text + "\n")
	results.pack()
	#results.configure(state="disabled")
	print(msg_text)

## The function that actually does data processing.
def formatfile(infilename, outfilename):
	f = open(infilename, "r")
	infile = csv.reader(f)
	
	## create outfile:
	outfilename = outfilename + '.csv'
	outfile = open(outfilename, "w")
	
	count = 0
	warning = 0
	linenum = 0
	for line in infile:
		linenum += 1
		
		## process header line
		if linenum == 1:
			messaging("Header: %s" % line)
			if len(line)==1:
				messaging("Error: First line (header) is not comma-separated. Check that the first line contains the headers. Also check that this is actually a CSV file.")
				return("error")
			header = line
			newheader = ','.join(map(str, (header[ i ] for i in range(0,len(header))))) + "\n"
			outfile.write( newheader )
			continue;
		
		## if this isn't a new person (the "office code" field is blank), go to next line.
		## If it is a new person, count a new person and go on to read the next 7(?) lines as well.
		#line[0] = line[0].strip('\"\'')
		if len(line[0])==0 :
			continue
		elif len(line[0])!=4:
			messaging("Warning: On line %s, Office Code may be incorrect: %s" % (linenum, line[0]) )
		count += 1
		
		## tax form fields are on the second line
		next(infile)
		next(infile)
		messaging("Line 2: %s" % line)
		# line2 = line2.split(",")
		# line2[-1] = line2[-1].strip('\"\'')
		# if len(header) != len(line2):
			# messaging("Error: Different number of columns in record %s, line 2 (file line # %s) than in header." % (count, linenum) )
			# messaging("Person: %s\nProblematic line: %s" % (line,line2) )
			# return("error")
		# line2[-1] = line2[-1].strip()
		# #get the index # for these columns: futureproofing, in case more cols are added later
		# tax_headers = ['w4', 'witholding', 'additional amount', 'exempt']
		# for i in range(len(tax_headers)):
			# try:
				# #index = header.index( tax_headers[i] )
				# #line[index] = line2[index]
				# #messaging("Step %s, index of %s: %s" % ( i, tax_headers[i] , header.index( tax_headers[i] )) )
				# line[ header.index( tax_headers[i] ) ] = line2[ header.index( tax_headers[i] ) ]
			# except ValueError:
				# if count==1:
					# messaging("Warning: %s column is missing" % tax_headers[i] )
					# if warning == 0:
						# messaging("File Headers: %s" % header)
						# warning = 1
			# except IndexError:
				# messaging("Error on line %s of file. Count: %s, Step: %s, Index of %s: %s" % ( linenum, count, i, tax_headers[i] , header.index( tax_headers[i] )) )
				# messaging("Person: %s\nProblematic line: %s" % (line,line2) )
				# return("error")
		
		# # effective date is on the third line
		# line3 = infile.readline()
		# line3 = line3.split(",")
		# if len(header) != len(line2):
			# messaging("Error: Different number of columns in record %s, line 3 than in header" % count)
			# return("error")
		# line3[-1] = line3[-1].strip()
		# try:
			# effe_index = header.index("effective date")
			# line[effe_index] = line3[effe_index]
		# except ValueError:
			# if count==1:
				# messaging("Warning: effective date column is missing" )
				# if warning == 0:
					# messaging("File Headers: %s" % header)
					# warning = 1
		
		# # skip 2 lines b/c they add nothing
		# temp = infile.readline() #line 4
		# temp = infile.readline() #line 5
		
		# # EEOC categories are on the sixth line
		# line6 = infile.readline()
		# line6 = line6.split(",")
		# line6[-1] = line6[-1].strip()
		# if len(header) != len(line2):
			# messaging("Error: Different number of columns in record %s, line 6 than in header" % count)
			# return("error")
		# eeoc_headers = ["eeoc", "eeoc 2"]
		# for i in range(2):
			# try:
				# line[ header.index( eeoc_headers[i] ) ] = line2[ header.index( eeoc_headers[i] ) ]
			# except ValueError:
				# if count==1:
					# messaging("Warning: %s column is missing" % eeoc_headers[i] )
					# if warning == 0:
						# messaging("File Headers: %s" % header)
						# warning = 1
		
		# #write our consolidated data to file!
		# linestr = ','.join(map(str, (line[ i ] for i in range(0,len(header)) ))) + "\n"
		# # http://stackoverflow.com/questions/44778/
		# outfile.write( linestr )
	
	outfile.close()
	f.close()
	
	messaging( "Finished %s! %s records processed." % (outfilename, count) )
	return("ok")

## The function that runs when you click the button. Filepicker lives here.
def selectfile():
	# https://stackoverflow.com/questions/39950322/tkinter-filedialog-reading-in-a-file-path
	infilename = tkfile.askopenfilename(parent=root, filetypes=myFormats, title='Choose a file')
	if len(infilename) == 0:
		messaging("Error: No file was selected, try again.\n")
	elif infilename != None:
		messaging("The file %s has been selected" % infilename)
		
		#run formatfile(), including what the output should be named
		was_error = formatfile(infilename, outname.get() )
		
		if(was_error=="ok"):
			folderpath = '/'.join(infilename.split('/')[0:-1])
			messaging("Saved in folder " + folderpath + "\n")
			#system("start " + folderpath) # this just opens a command line, weirdly.
			#subprocess.Popen('explorer "' + folderpath + '"') # This does not wanna work.
		elif(was_error=="error"):
			return

## BEGIN THE GUI PART (WOO!)
root = tk.Tk()
#root.withdraw()  # if we don't want a full GUI (want to run the whole thing from command line, but still have the filepicker prompt), keep the root window from appearing
root.title("Consolidate DocuSign Exports")

tk.Label(root, text="Welcome to this very exciting Python program!").pack()

fileimport_grp = tk.LabelFrame(root, text="Choose File", padx=5, pady=5)
fileimport_grp.pack(padx=10, pady=10)

tk.Label(fileimport_grp, text="What to name the output file? (Don't include .csv at the end.)", anchor="w").pack(fill="x")
tk.Label(fileimport_grp, text="Note: if that file already exists, it will be overwritten!", anchor="w").pack(fill="x")

outname = tk.Entry(fileimport_grp)
outname.pack(padx=5, pady=5)
outname.focus_set()
outname.insert(0, "output")
# out = outname.get() # nope this is a global variable so don't need to worry about passing it to fn I guess??

b = tk.Button(fileimport_grp, text="Choose a CSV file", command=selectfile , padx=2, pady=5)
b.pack(padx=10, pady=10)

results = tk.Text(root, height=10, width=50)
results.config(font="Helvetica")
results.pack()
#results.configure(state="disabled")

root.mainloop()

outname = tk.Entry(fileimport_grp, width=25)
outname.pack()
#weird that you have to make the Entry thing twice, once before and once after calling mainloop, but that's what they do here:
# http://effbot.org/tkinterbook/entry.htm

## on making this into a Windows executable:
# py2exe doesn't like it, so ugh.