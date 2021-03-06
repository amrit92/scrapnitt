'''
NITT Result Scraper


@author : Amrit Sahoo (amritsahoo@gmail.com), CSE, NIT Trichy, 2010-14
@author : Billa Prashanth Reddy (prashanthreddybilla@gmail.com), CSE, NIT Trichy, 2010-14


Copyright (C) 2014 Amrit Sahoo, Billa Prashanth Reddy 
Everyone is permitted to copy and distribute verbatim copies of this license document, but changing it is not allowed.

'''

import subprocess
import xlwt
import os
import string
import mechanize
import xlutils
import cookielib
import xlrd
from xlutils.copy import copy
import sys
from Tkinter import *
import tkMessageBox
import ttk
complete = 1
def get_result(newvalue, sem, dept, year):
	br = mechanize.Browser()
	br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
	br.open('http://nitt.edu/prm/nitreg/ShowRes.aspx')
	br.select_form("Form1")
	br["TextBox1"]= newvalue
	response = br.submit()
	br.select_form("Form1")
	br.set_all_readonly(False)
	#br["__EVENTTARGET"] = 'Dt1'
	try:
		control = br.form.find_control("Dt1")
		flag = 1
		for ij in control.items:
			if((str(sem) != str(ij.name))):
				continue
			elif(str(sem) == str(ij.name)):
				br.select_form("Form1")
				#br.set_all_readonly(False)
				br["__EVENTTARGET"] = 'Dt1'
				br["Dt1"]=[sem]
				br.find_control("Button1").disabled = True
				output = br.submit().read()
				count = int(newvalue) - int(dept+year+"000")
				regex_gpa = "id=\"LblGPA\""
				regex_rollnum = "id=\"LblEnrollmentNo\""
				regex_name = "id=\"LblName\""
				fileio = open("Result/OUTPUT.doc","a")
				workbook = xlrd.open_workbook('Result/result.xls')
				mysheet_read = workbook.sheet_by_index(0)
				mysheet_write = copy(workbook)
				mysheet = mysheet_write.get_sheet(0)
				mysheet.write(0, 0, "Name")
				mysheet.write(0, 1, "Roll number")
				mysheet.write(0, 2, "Gpa")
				tempr = open("temporary_files/tempfile1.txt","r")
				temp_gpar = open("temporary_files/tempfile2.txt","r")
				for line in string.split(output, '\n'):
					if regex_rollnum in line:
						text1 = "  Roll number  : " + ((line.split(">")[3]).split("<"))[0]
						fileio.write(text1)
						fileio.write("\n")
						mysheet.write(count, 1, ((line.split(">")[3]).split("<"))[0])
					elif regex_name in line:
						text2 = "  Name  : " + ((line.split(">")[3]).split("<"))[0]
						fileio.write(text2)
						fileio.write("\n")
						mysheet.write(count, 0, ((line.split(">")[3]).split("<"))[0])
					elif regex_gpa in line:
						text3 = "  GPA  : " + ((line.split(">")[3]).split("<"))[0]
						fileio.write(text3)
						fileio.write("\n")
						fileio.write("\n")
						mysheet.write(count, 2, ((line.split(">")[3]).split("<"))[0])
						v1 = str(int(tempr.readline())+1)
						tempr.close()
						tempw = open("temporary_files/tempfile1.txt","w")
						tempw.write(v1)
						tempw.close()
						cur_gpa = float(temp_gpar.readline())+float(((line.split(">")[3]).split("<"))[0])
						temp_gpar.close()
						temp_gpaw = open("temporary_files/tempfile2.txt","w")
						temp_gpaw.write(str(cur_gpa))
						temp_gpaw.close()
						if(float(((line.split(">")[3]).split("<"))[0]) == 10.00):
							t1 = open("temporary_files/ten.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/ten.txt","w")
							try:
								t11 = int(t1value) + 1
						
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
						elif(float(((line.split(">")[3]).split("<"))[0]) >=9.00):
							t1 = open("temporary_files/nine.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/nine.txt")
							try:
								t11 = int(t1value) + 1
						
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
						elif(float(((line.split(">")[3]).split("<"))[0]) >=8.00):
							t1 = open("temporary_files/eight.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/eight.txt","w")
							try:
								t11 = int(t1value) + 1
						
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
						elif(float(((line.split(">")[3]).split("<"))[0]) >=7.00):
							t1 = open("temporary_files/seven.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/seven.txt","w")
							try:
								t11 = int(t1value) + 1
		
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
						elif(float(((line.split(">")[3]).split("<"))[0]) >=6.00):
							t1 = open("temporary_files/six.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/six.txt","w")
							try:
								t11 = int(t1value) + 1
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
						elif(float(((line.split(">")[3]).split("<"))[0]) >=5.00):
							t1 = open("temporary_files/five.txt","r")
							t1value = t1.readline()
							t1.close()
							t2 = open("temporary_files/five.txt","w")
							try:
								t11 = int(t1value) + 1
								t2.write(str(t11))
								t2.close()
							except ValueError:
								pass
				mysheet_write.save("Result/result.xls")
		return 1
	except mechanize.ControlNotFoundError:
		return 1
	except URLError:
		return -1
	
	

def main_function(dept, year, sem):
	
	if not os.path.exists("temporary_files"):
    		os.makedirs("temporary_files")
    	if not os.path.exists("Result"):
    		os.makedirs("Result")
	global complete
	complete = 0
	value = dept+year+"000"
	book = xlwt.Workbook()
	sheet1 = book.add_sheet("excelsheet")
	col1 = sheet1.col(0)
	col1.width = 256*40
	book.save("Result/result.xls")
	temp1 = open("temporary_files/tempfile1.txt","w")
	temp1.write("0")
	temp1.close()
	temp2 = open("temporary_files/tempfile2.txt","w")
	temp2.write("0")
	temp2.close()
	temp10 = open("temporary_files/ten.txt","w")
	temp10.write("0")
	temp10.close()
	temp9 = open("temporary_files/nine.txt","w")
	temp9.write("0")
	temp9.close()
	temp8 = open("temporary_files/eight.txt","w")
	temp8.write("0")
	temp8.close()
	temp7 = open("temporary_files/seven.txt","w")
	temp7.write("0")
	temp7.close()
	temp6 = open("temporary_files/six.txt","w")
	temp6.write("0")
	temp6.close()
	temp5 = open("temporary_files/five.txt","w")
	temp5.write("0")
	temp5.close()
	for j in range(1,107):
        	newvalue = str(int(value) + j)
        	returnval = get_result(newvalue,sem,dept,year)
        	if(returnval == 1):
        		continue;
        	while(returnval == -1):
        		returnval = get_result(newvalue,sem,dept,year)
	temp1 = open("temporary_files/tempfile1.txt","r")
	temp2 = open("temporary_files/tempfile2.txt","r")
	no = float(temp1.readline())
	total = float(temp2.readline())
	avg = total/no
	fileavg = open("Result/Statistics.doc","w")
	fileavg.write("Department code :" + dept)
	
	fileavg.write("\n\n")
	fileavg.write("Year : "+ year)
	fileavg.write("\n\n")
	if(str(var3.get()) == "1"):
		fileavg.write("Odd semester")
	if(str(var3.get()) == "2"):
		fileavg.write("Even Semester")
	fileavg.write("\n\n")
	fileavg.write("Average GPA = "+str(avg))
	fileavg.write("\n\n")
	temp10 = open("temporary_files/ten.txt","r")
	temp9 = open("temporary_files/nine.txt","r")
	temp8 = open("temporary_files/eight.txt","r")
	temp7 = open("temporary_files/seven.txt","r")
	temp6 = open("temporary_files/six.txt","r")
	temp5 = open("temporary_files/five.txt","r")
	fileavg.write("No. of 10 pointers : "+temp10.readline())
	fileavg.write("\n")
	fileavg.write("No. of 9 pointers : "+temp9.readline())
	fileavg.write("\n")
	fileavg.write("No. of 8 pointers : "+temp8.readline())
	fileavg.write("\n")
	fileavg.write("No. of 7 pointers : "+temp7.readline())
	fileavg.write("\n")
	fileavg.write("No. of 6 pointers : "+temp6.readline())
	fileavg.write("\n")
	fileavg.write("No. of 5 pointers : "+temp5.readline())
	fileavg.close()
	temp10.close()
	temp9.close()
	temp8.close()
	temp7.close()
	temp6.close()
	temp5.close()
	tkMessageBox.showinfo("Complete :)",'Check out the result folder.')
	var1.set("status : completed")
	os.remove("temporary_files/tempfile1.txt")
	os.remove("temporary_files/tempfile2.txt")
	os.remove("temporary_files/ten.txt")
	os.remove("temporary_files/nine.txt")
	os.remove("temporary_files/eight.txt")
	os.remove("temporary_files/seven.txt")
	os.remove("temporary_files/six.txt")
	os.remove("temporary_files/five.txt")
	os.rmdir("temporary_files")
	complete = 1
	frame.destroy()
	
def call_function(value):
	if(value == "Select department"):
		tkMessageBox.showinfo("OOps","Select your department")
		frame.destroy()
		subprocess.call("python run.py", shell=True)
	
	year = "110"
	sem = ""
	var1.set("status : started")
	if(str(var2.get()) == "2013"):
		year = "110"
	if(str(var2.get()) == "2012"):
		year = "109"
	if(str(var3.get()) == "1" and year == "109" ):
		sem = "84"
	if(str(var3.get()) == "2" and year == "109" ):
		sem = "77"
	if(str(var3.get()) == "1" and year == "110" ):
		tkMessageBox.showinfo("Not yet", "results are not yet out")
		frame.destroy()
		subprocess.call("python run.py", shell=True)
	if(str(var3.get()) == "2" and year == "110" ):
		sem = "88"
	tkMessageBox.showinfo("Started","Click ok to start. You will be notified once the process is completed.")
	if(value == "Architecture"):
		main_function("101",year,sem);
	elif(value == "Chemical"):
		main_function("102",year,sem);
	elif(value == "Civil"):
		main_function("103",year,sem);
	elif(value == "CSE"):
		main_function("106",year,sem);
	elif(value == "EEE"):
		main_function("107",year,sem);
	elif(value == "ECE"):
		main_function("108",year,sem);
	elif(value == "ICE"):
		main_function("110",year,sem);
	elif(value == "Mechanical"):
		main_function("111",year,sem);
	elif(value == "Production"):
		main_function("114",year,sem);
	elif(value == "Metallurgy"):
		main_function("112",year,sem);
	elif(value == "MCA"):
		main_function("205",year,sem);
	
#####################################gui
  
def newcommand():
	frame.destroy()
	subprocess.call("python run.py", shell=True)
def showl():
	tkMessageBox.showinfo("GNU-GPL","This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.")
def helpcontent():
	content = "Select your department, current year, and semester to get the Grade point average of all the students in the class. The results are saved in a doc file as well as an excel file in the current folder of the exe file. Also statistics.doc is created at the end of the process to give more information about the class performance. Incase of any queries, please contact (amritsahoo@gmail.com) / (prashanthreddybilla@gmail.com)"
	tkMessageBox.showinfo("About", content)
def showabout():
	tkMessageBox.showinfo("About", "(c)2014 GNU-GPL License, Authors - Amrit sahoo (amritsahoo@gmail.com), Billa Prashanth Reddy (prashanthreddybilla@gmail.com)")
def displayOption():
	
	button.config(state='disabled')
	call_function(optionMenuWidget.cget("text"))
   
def quitapp():
	if(complete == 0):
		if tkMessageBox.askyesno("quit",'Project is not saved. Ignore changes and quit?'):
			frame.destroy()
	else:
		frame.destroy()
	
frame = Tk()
DEFAULTVALUE_OPTION = "Select department"    
    
frame.title("Nitt result scraper")
frame["padx"] = 60
frame["pady"] = 40       
frame.wm_iconbitmap(bitmap = "nitt.ico")
photo = PhotoImage(file="nitt.gif")
w = Label(frame, image=photo)
w.photo = photo
w.pack()
optionFrame = Frame(frame)
    
optionLabel = Label(optionFrame)
optionLabel["text"] = DEFAULTVALUE_OPTION
optionLabel.pack(side=LEFT)

   
optionTuple = ("Architecture", "Chemical", "Civil", "CSE", "EEE", "ECE", "ICE", "Mechanical", "Metallurgy", "Production", "MCA")

var2 = StringVar(frame)
var2.set("2013")
sb = Spinbox(frame, from_=2012, to=2013, textvariable=var2)
sb.pack()

var3 = IntVar()
R1 = Radiobutton(frame, text="Odd semester", variable=var3, value=1)
R1.pack( anchor = W )
R1.select()
R2 = Radiobutton(frame, text="Even semester", variable=var3, value=2)
R2.pack( anchor = W )


defaultOption = StringVar()
optionMenuWidget = apply(OptionMenu, (optionFrame, defaultOption) + optionTuple)
defaultOption.set(DEFAULTVALUE_OPTION)
optionMenuWidget["width"] = 15
optionMenuWidget.pack(side=LEFT)

optionFrame.pack()

button = Button(frame, text="Submit", command=displayOption)
button.pack() 
var1 = StringVar()
label = Label(frame, textvariable=var1, relief=RAISED )

var1.set("status : Not started")
label.pack()
menubar = Menu(frame)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=newcommand)
filemenu.add_separator()

filemenu.add_command(label="Exit", command=quitapp)
menubar.add_cascade(label="File", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=helpcontent)
helpmenu.add_command(label="About", command=showabout)
helpmenu.add_command(label="License", command=showl)
menubar.add_cascade(label="Help", menu=helpmenu)

frame.config(menu=menubar)
    
frame.mainloop()
########################################3

