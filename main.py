

import mechanize
import cookielib

# Browser
br = mechanize.Browser()
br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
br.open('http://nitt.edu/prm/nitreg/ShowRes.aspx')
br.select_form("Form1")

br["TextBox1"]="106110005"
response = br.submit()
#print response.read()
br.select_form("Form1")
br.set_all_readonly(False)
br["__EVENTTARGET"] = 'Dt1'
control = br.form.find_control("Dt1")
a = 0
for i in control.items:
	br.select_form("Form1")
	br.set_all_readonly(False)
	br["__EVENTTARGET"] = 'Dt1'
	br["Dt1"]=[control.items[a].name]
	br.find_control("Button1").disabled = True
	#print br.submit().read()
	fout = open('tmp%d.html' %( a + 1 ), 'w')
	fout.writelines(br.submit().read())
	fout.close()
	a = a + 1



