# Code to Install pyPDF module using terminal
# sudo python -m pip install pyPdf

#!/usr/bin/env python
import copy, sys, os
from pyPdf import PdfFileWriter, PdfFileReader

def remove_blank_pages(filename):
	reader = PdfFileReader(open(filename))
	writer = PdfFileWriter()
	NUM_PAGES = reader.getNumPages()

	for i in range(NUM_PAGES):
		p = reader.getPage(i)
		if i < NUM_PAGES - 2:
			writer.addPage(p)

	new_pdf = open("new_pdfs/" + filename, 'w')
	writer.write(new_pdf)
	new_pdf.close()

# Get current working directory
cwd = os.getcwd()

# Create empty directory to store new pdfs
ned = cwd + "/new_pdfs/"
if not os.path.exists(ned):
	os.makedirs(ned)

for filename in os.listdir(cwd):
    if filename.endswith(".pdf"):
    	remove_blank_pages(filename)


