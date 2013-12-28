import mechanize
import cookielib
import sys
import subprocess
import xlwt
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

regex_gpa = "id=\"LblGPA\""
regex_rollnum = "id=\"LblEnrollmentNo\""
regex_name = "id=\"LblName\""
writefile = open("outputfile","w")
fileio = open("OUTPUT.xls","a")
writefile.write(output)
writefile.close()
textfile = open("outputfile","r")
book = xlwt.Workbook()
sheet1 = book.add_sheet("sheetmy")
sheet1.write(0, 0, "Name")
sheet1.write(0, 1, "Roll number")
sheet1.write(0, 2, "Gpa")
book.save("trial.xls")
for line in textfile.readlines():
	if regex_rollnum in line:
		text1 = "  Roll number  : " + ((line.split(">")[3]).split("<"))[0]
		fileio.write(text1)
		fileio.write("\n")
	elif regex_name in line:
		text2 = "  Name  : " + ((line.split(">")[3]).split("<"))[0]
		fileio.write(text2)
		fileio.write("\n")
	elif regex_gpa in line:
		text3 = "  GPA  : " + ((line.split(">")[3]).split("<"))[0]
		fileio.write(text3)
		fileio.write("\n")
		fileio.write("\n")

