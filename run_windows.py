'''
NITT Result Scraper


@author : Amrit Sahoo (amritsahoo@gmail.com), CSE, NIT Trichy, 2010-14
@author : Prashanth Reddy Billa(prashanthreddybilla@gmail.com), CSE, NIT Trichy, 2010-14


Copyright (C) 2014 Amrit Sahoo, Billa Prashanth Reddy 
Everyone is permitted to copy and distribute verbatim copies of this license document, but changing it is not allowed.

'''
import thread
import subprocess
import xlwt
import os
from time import sleep
import string
import mechanize
import xlutils
import cookielib
import xlrd
import urllib2
from xlutils.copy import copy
import sys
from Tkinter import *
import tkMessageBox
import ttk
import socket
complete = 1
def get_result(newvalue, sem, dept, year, name, vyear, vsem):
	try:
		br = mechanize.Browser()
		br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
		br.open('http://nitt.edu/prm/nitreg/ShowRes.aspx',timeout=800.0)
		br.select_form("Form1")
		br["TextBox1"]= newvalue
		response = br.submit()
		br.select_form("Form1")
		br.set_all_readonly(False)
		error_code=""
	except urllib2.URLError, e:
		error_code =  (str(e.reason).split("]")[0]).split(" ")[1]
		if(error_code == "10060"):
			print "Temporary problem at server. Trying again..."
			sleep(5)
			return -1
		elif(error_code == "11004"):
			print "Problem with your internet connection. Please check your connection. Trying again..."
			sleep(5)
			return -1
		else:
			print "Problem at the Server. Trying again..."
			sleep(5)
			return -1
	except urllib2.HTTPError:
		#print "Server down. Trying again..."
		return 2
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
				fileio = open("Result/OUTPUT_"+name+"_"+vyear+"_"+vsem+".doc","a")
				workbook = xlrd.open_workbook("Result/result_"+name+"_"+vyear+"_"+vsem+".xls")
				mysheet_read = workbook.sheet_by_index(0)
				mysheet_write = copy(workbook)
				mysheet = mysheet_write.get_sheet(0)
				mysheet.write(0, 0, "Name")
				mysheet.write(0, 1, "Roll number")
				mysheet.write(0, 2, "Gpa")
				tempr = open("temporary_files/tempfile1.txt","r")
				temp_gpar = open("temporary_files/tempfile2.txt","r")
				myline = ""
				for line in string.split(output, '\n'):
					myline = line
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
						print v1 +" students completed successfully"
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
							t2 = open("temporary_files/nine.txt","w")
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
				mysheet_write.save("Result/result_"+name+"_"+vyear+"_"+vsem+".xls")
		return 1
	except mechanize.ControlNotFoundError:
		return 1
	except urllib2.URLError:
		print "Server refused the request. Trying again for the roll number : "+((myline.split(">")[3]).split("<"))[0]
		sleep(5)
		return -1
		
def main_function(dept, year, sem, name, vyear, vsem):
	
	if not os.path.exists("temporary_files"):
    		os.makedirs("temporary_files")
    	if not os.path.exists("Result"):
    		os.makedirs("Result")
	global complete
	complete = 0
	value = dept+year+"000"
	if os.path.exists("Result/result_"+name+"_"+vyear+"_"+vsem+".xls"):
		os.remove("Result/result_"+name+"_"+vyear+"_"+vsem+".xls")
	book = xlwt.Workbook()
	sheet1 = book.add_sheet("excelsheet")
	col1 = sheet1.col(0)
	col1.width = 256*40
	book.save("Result/result_"+name+"_"+vyear+"_"+vsem+".xls")
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
	for j in range(1,110):
        	newvalue = str(int(value) + j)
		print newvalue+" in progress"
        	returnval = get_result(newvalue,sem,dept,year,name,vyear,vsem)
        	if(returnval == 1):
        		continue;
			if(returnval == 2):
				print "Server is unavailable. Try again later"
				sys.exit(0)
        	while(returnval == -1):
				if(returnval == -1):
					returnval = get_result(newvalue,sem,dept,year,name,vyear,vsem)
	temp1 = open("temporary_files/tempfile1.txt","r")
	temp2 = open("temporary_files/tempfile2.txt","r")
	no = float(temp1.readline())
	total = float(temp2.readline())
	avg = total/no
	fileavg = open("Result/statistics_"+name+"_"+vyear+"_"+vsem+".doc","w")
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
	print "Process complete! Check out the Result folder"
	print "Thank You for using the App"
	tkMessageBox.showinfo("Complete :)",'Check out the result folder.')
	var1.set("status : completed")
	complete = 1
	frame.destroy()
	
def call_function(value):
	if(value == "Select department"):
		tkMessageBox.showinfo("OOps","Select your department")
		frame.destroy()
		subprocess.call("run.exe", shell=True)
	semname = ""
	if(str(var3.get()) == "1"):
		semname = "Odd Semester"
	if(str(var3.get()) == "2"):
		semname = "Even Semester"
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
		subprocess.call("run.exe", shell=True)
	if(str(var3.get()) == "2" and year == "110" ):
		sem = "88"
	tkMessageBox.showinfo("Process about to start","Click ok to start. You will be notified once the process is completed. Note: The process will run for 110 students by default as the exact number of students in a class is not accurately known.")
	if(value == "Architecture"):
		main_function("101",year,sem,value,str(var2.get()),semname);
	elif(value == "Chemical"):
		main_function("102",year,sem,value,str(var2.get()),semname);
	elif(value == "Civil"):
		main_function("103",year,sem,value,str(var2.get()),semname);
	elif(value == "CSE"):
		main_function("106",year,sem,value,str(var2.get()),semname);
	elif(value == "EEE"):
		main_function("107",year,sem,value,str(var2.get()),semname);
	elif(value == "ECE"):
		main_function("108",year,sem,value,str(var2.get()),semname);
	elif(value == "ICE"):
		main_function("110",year,sem,value,str(var2.get()),semname);
	elif(value == "Mechanical"):
		main_function("111",year,sem,value,str(var2.get()),semname);
	elif(value == "Production"):
		main_function("114",year,sem,value,str(var2.get()),semname);
	elif(value == "Metallurgy"):
		main_function("112",year,sem,value,str(var2.get()),semname);
	elif(value == "MCA"):
		main_function("205",year,sem,value,str(var2.get()),semname);
	
#####################################gui
  
def newcommand():
	if(complete == 0):
		if tkMessageBox.askyesno("New",'Process is not complete. Start a new process?'):
			frame.destroy()
	else:
		frame.destroy()
		subprocess.call("run.exe", shell=True)
	
def showl():
	tkMessageBox.showinfo("GNU-GPL","This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.")
def helpcontent():
	content = "Select your department, current year, and semester to get the Grade point average of all the students in the class. The results are saved in a doc file as well as an excel file in the Results folder. Also statistics.doc is created at the end of the process to give more information about the class performance. Incase of any queries, please contact (amritsahoo@gmail.com) / (prashanthreddybilla@gmail.com)"
	tkMessageBox.showinfo("About", content)
def showabout():
	tkMessageBox.showinfo("About", "(c)2014 GNU-GPL License, Authors - Amrit sahoo (amritsahoo@gmail.com), Prashanth Reddy Billa(prashanthreddybilla@gmail.com)")
def displayOption():
	
	button.config(state='disabled')
	call_function(optionMenuWidget.cget("text"))
   
def quitapp():
	if(complete == 0):
		if tkMessageBox.askyesno("quit",'Process is not complete. Quit?'):
			frame.destroy()
	else:
		frame.destroy()


	
frame = Tk()
DEFAULTVALUE_OPTION = "Select department"    
    
frame.title("Nitt result scraper")
frame["padx"] = 60
frame["pady"] = 40       
frame.iconbitmap(default='nitt.ico')
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

