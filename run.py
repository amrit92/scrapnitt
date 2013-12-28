import subprocess
import xlwt
value = 106110000
book = xlwt.Workbook()
sheet1 = book.add_sheet("excelsheet")
col1 = sheet1.col(0)
col1.width = 256*40
book.save("result.xls")
for j in xrange(1,107):
	newvalue = str(value + j)
	subprocess.call(" python get_result.py "+newvalue, shell=True)
	
