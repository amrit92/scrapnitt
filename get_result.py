import mechanize
import xlutils
import cookielib
import sys
import subprocess
import xlwt
import xlrd
from xlutils.copy import copy
# Browser
br = mechanize.Browser()

br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]


br.open('http://nitt.edu/prm/nitreg/ShowRes.aspx')



br.select_form("Form1")


br["TextBox1"]= sys.argv[1]
response = br.submit()
#print response.read()

#(contain.split(">")[1]).split("<")[0];
br.select_form("Form1")
br.set_all_readonly(False)
br["__EVENTTARGET"] = 'Dt1'
br["Dt1"]=["88"]
br.find_control("Button1").disabled = True
output = br.submit().read()
count = int(sys.argv[1]) - int("106110000")
regex_gpa = "id=\"LblGPA\""
regex_rollnum = "id=\"LblEnrollmentNo\""
regex_name = "id=\"LblName\""
writefile = open("outputfile","w")
fileio = open("OUTPUT.doc","a")
writefile.write(output)
writefile.close()
textfile = open("outputfile","r")
workbook = xlrd.open_workbook('result.xls')
mysheet_read = workbook.sheet_by_index(0)
mysheet_write = copy(workbook)
mysheet = mysheet_write.get_sheet(0)
mysheet.write(0, 0, "Name")
mysheet.write(0, 1, "Roll number")
mysheet.write(0, 2, "Gpa")
for line in textfile.readlines():
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
mysheet_write.save("result.xls")

