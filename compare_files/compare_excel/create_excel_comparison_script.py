from os import listdir
from os import path

wd25_rpts = listdir("wd25_rpts/")
wd26_rpts = listdir("wd26_rpts/")

NUM_FILES = len(wd25_rpts)

f = open("compare_rpts.sh","w")
for i in range(NUM_FILES):
	excel_cmp = "bash /Users/josefnunez/ExcelCompare/bin/excel_cmp"
	rpt_name = "'" + wd25_rpts[i].replace("WD25 - PROD - ","") + "'"
	rpt_wd25 = "'" + "wd25_rpts/" + wd25_rpts[i] + "'"
	rpt_wd26 = "'" + "wd26_rpts/" + wd26_rpts[i] + "'"
	output = "results/" + rpt_name
	# Place filenames between single quotes (') to prevent
	# runtime errors in terminal
	f.write(excel_cmp + " " + rpt_wd25 + " " + rpt_wd26 + " > " + output + "\n")
	f.write("echo" + " " + rpt_name + " " + "comparison complete\n")
f.close()

results = listdir("results/")
f = open("summary_results.txt","w")
for r in results:
	curr = open("results/" + r, "r")
	lines = curr.readlines()
	last = lines[-1] if len(lines) > 0 else r + " " + "comparison ERROR\n"
	f.write(last)
	curr.close()
f.close()