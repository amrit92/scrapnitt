import subprocess
import xlwt
import os
import string
import mechanize
import xlutils
import cookielib
import xlrd
from xlutils.copy import copy

def get_result(newvalue):
	br = mechanize.Browser()
	br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
	br.open('http://www.nitt.edu/prm/nitreg/ShowRes.aspx')
	br.select_form("Form1")
	br["TextBox1"]= newvalue
	response = br.submit()
	br.select_form("Form1")
	br.set_all_readonly(False)
	br["__EVENTTARGET"] = 'Dt1'
	br["Dt1"]=["88"]
	br.find_control("Button1").disabled = True
	output = br.submit().read()
	count = int(newvalue) - int("106110000")
	regex_gpa = "id=\"LblGPA\""
	regex_rollnum = "id=\"LblEnrollmentNo\""
	regex_name = "id=\"LblName\""
	fileio = open("OUTPUT.doc","a")
	workbook = xlrd.open_workbook('result.xls')
	mysheet_read = workbook.sheet_by_index(0)
	mysheet_write = copy(workbook)
	mysheet = mysheet_write.get_sheet(0)
	mysheet.write(0, 0, "Name")
	mysheet.write(0, 1, "Roll number")
	mysheet.write(0, 2, "Gpa")
	tempr = open("tempfile1","r")
	temp_gpar = open("tempfile2","r")
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
                	tempw = open("tempfile1","w")
                	tempw.write(v1)
                	tempw.close()
                	cur_gpa = float(temp_gpar.readline())+float(((line.split(">")[3]).split("<"))[0])
                	temp_gpar.close()
                	temp_gpaw = open("tempfile2","w")
                	temp_gpaw.write(str(cur_gpa))
                	temp_gpaw.close()
                	if(float(((line.split(">")[3]).split("<"))[0]) == 10.00):
                		t1 = open("ten","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("ten","w")
                		try:
                			t11 = int(t1value) + 1
                		
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
                	elif(float(((line.split(">")[3]).split("<"))[0]) >=9.00):
                		t1 = open("nine","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("nine","w")
                		try:
                			t11 = int(t1value) + 1
                		
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
                	elif(float(((line.split(">")[3]).split("<"))[0]) >=8.00):
                		t1 = open("eight","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("eight","w")
                		try:
                			t11 = int(t1value) + 1
                		
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
                	elif(float(((line.split(">")[3]).split("<"))[0]) >=7.00):
                		t1 = open("seven","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("seven","w")
                		try:
                			t11 = int(t1value) + 1
        
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
                	elif(float(((line.split(">")[3]).split("<"))[0]) >=6.00):
                		t1 = open("six","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("six","w")
                		try:
                			t11 = int(t1value) + 1
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
                	elif(float(((line.split(">")[3]).split("<"))[0]) >=5.00):
                		t1 = open("five","r")
                		t1value = t1.readline()
                		t1.close()
                		t2 = open("five","w")
                		try:
                			t11 = int(t1value) + 1
                			t2.write(str(t11))
                			t2.close()
                		except ValueError:
                			pass
	mysheet_write.save("result.xls")



value = 106110000
book = xlwt.Workbook()
sheet1 = book.add_sheet("excelsheet")
col1 = sheet1.col(0)
col1.width = 256*40
book.save("result.xls")
temp1 = open("tempfile1","w")
temp1.write("0")
temp1.close()
temp2 = open("tempfile2","w")
temp2.write("0")
temp2.close()
temp10 = open("ten","w")
temp10.write("0")
temp10.close()
temp9 = open("nine","w")
temp9.write("0")
temp9.close()
temp8 = open("eight","w")
temp8.write("0")
temp8.close()
temp7 = open("seven","w")
temp7.write("0")
temp7.close()
temp6 = open("six","w")
temp6.write("0")
temp6.close()
temp5 = open("five","w")
temp5.write("0")
temp5.close()
for j in xrange(1,107):
        newvalue = str(value + j)
        get_result(newvalue)
temp1 = open("tempfile1","r")
temp2 = open("tempfile2","r")
no = float(temp1.readline())
total = float(temp2.readline())
avg = total/no
fileavg = open("Statistics.doc","w")
fileavg.write("Average GPA = "+str(avg))
fileavg.write("\n\n")
temp10 = open("ten","r")
temp9 = open("nine","r")
temp8 = open("eight","r")
temp7 = open("seven","r")
temp6 = open("six","r")
temp5 = open("five","r")
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
os.remove("tempfile1")
os.remove("tempfile2")
os.remove("ten")
os.remove("nine")
os.remove("eight")
os.remove("seven")
os.remove("six")
os.remove("five")
