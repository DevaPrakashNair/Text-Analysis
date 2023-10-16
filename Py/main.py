# import module
import openpyxl

import text_analysis
import text_extraction

# load excel with its path
wrkbk = openpyxl.load_workbook(r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\Input.xlsx")

sh = wrkbk.active
d ={}
# iterate through excel and display data
for row in sh.iter_rows(min_row=2, min_col=1):
	if str(row[0].value)[-1] == str(0):
		d[int(row[0].value)] = row[1].value
	else:
		d[row[0].value] = row[1].value
print(d)

# extract text from online article - run the codes in line 17 to perform text extraction
result = text_extraction.text_extraction(d)
# print(result)

# text analysis
text_analysis.text_analysis()
