import csv, openpyxl
from openpyxl import Workbook
#for an exported Ct csv
#https://www.reddit.com/r/learnpython/comments/63lvf7/really_confused_openpyxl_keeps_overwriting_my/

book = openpyxl.load_workbook("sample.xlsx")
sheet = book.active
count = 12
count2 = 12
count3 = 12
count4 = 12

book.save("sample.xlsx")


with open('first.csv') as csvfile:
	readCSV = csv.reader(csvfile,delimiter=',')
	for row in readCSV:
		#print (row[1])
		if row == "":
			print ("next")
		else:
			#print (row[0] + ",",row[1]+",",row[2])
			#row[1] is where the name is, row[2] has the Ct
			if "1e1" in row[1].lower():
				print ("1e1 found " + row[1])
				while sheet.cell(row=count,column=5).value != None:
					count += 1
				sheet.cell(row=count,column=5).value = row[2]			

			elif "1e2" in row[1].lower():
				print ("1e2 found " + row[2])
				while sheet.cell(row=count2,column=6).value != None:
					count2 += 1
				sheet.cell(row=count2,column=6).value = row[2]	

			elif "1e3" in row[1].lower():
				print ("1e3 found " + row[2])
				while sheet.cell(row=count3,column=7).value != None:
					count3 += 1
				sheet.cell(row=count3,column=7).value = row[2]	

			elif "1e4" in row[1].lower():
				print ("1e4 found " + row[2])
				while sheet.cell(row=count4,column=8).value != None:
					count4 += 1
				sheet.cell(row=count4,column=8).value = row[2]	

	book.save("sample.xlsx")


			
'''name = row[1].split()
for temp in name:
print (temp)'''
