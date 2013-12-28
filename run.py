import subprocess
value = 106110000
for j in xrange(1,107):
	newvalue = str(value + j)
	subprocess.call(" python get_result.py "+newvalue, shell=True)
	
