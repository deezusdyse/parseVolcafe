import xlwt
import os
from xlrd import open_workbook
from xlutils.copy import copy
from dateutil import parser
import time
import datetime

## Converts relevant sections in pdf to text
def pdf_to_txt(filename, separator, threshold):
	from cStringIO import StringIO
	from pdfminer.converter import LTChar, TextConverter
	from pdfminer.layout import LAParams
	from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
	from pdfminer.pdfpage import PDFPage

	class CsvConverter(TextConverter):
		def __init__(self, *args, **kwargs):
			TextConverter.__init__(self, *args, **kwargs)
			self.separator = separator
			self.threshold = threshold

		def end_page(self, i):
			from collections import defaultdict
			lines = defaultdict(lambda: {})
			for child in self.cur_item._objs: 
				if isinstance(child, LTChar):
					(_, _, x, y) = child.bbox
					line = lines[int(-y)]
					line[x] = child._text.encode(self.codec) 
			
			# Identifies transitions in text with keywords
			p = False
			vol = 0
			for y in sorted(lines.keys()):

				line = lines[y]

				if p == True:
					self.outfp.write(self.line_creator(line))
					self.outfp.write("\n")
									
				if "ORIGIN DIFFERENTIALS" in self.line_creator(line):
					p = True
				
					self.outfp.write(self.line_creator(line))
					self.outfp.write("\n")
					
				if "WEEKLY MARKET REVIEW" in self.line_creator(line):
					self.outfp.write(self.line_creator(line))
					self.outfp.write("\n")

				if "volcafe ltd" in self.line_creator(line).strip().lower():
					vol+=1
				
				if vol==1 and "weekly" not in self.line_creator(line).strip().lower():
					if (self.line_creator(line).strip().strip() != "") and "201" in self.line_creator(line).split(",")[-1]:
							self.line_creator(line)
							self.outfp.write(self.line_creator(line))
							self.outfp.write("\n")   	
							vol += 1

		def line_creator(self, line):
			keys = sorted(line.keys())
			average_distance = sum([keys[i] - keys[i - 1] for i in range(1, len(keys))]) / len(keys)
			result = [line[keys[0]]]
			for i in range(1, len(keys)):
				if (keys[i] - keys[i - 1]) > average_distance * self.threshold:
					result.append(self.separator)
				result.append(line[keys[i]])
			printable_line = ''.join(result)
			return printable_line

	rsrc = PDFResourceManager()
	outfp = StringIO()
	device = CsvConverter(rsrc, outfp, codec="utf-8", laparams=LAParams())

	fp = open(filename, 'rb')

	interpreter = PDFPageInterpreter(rsrc, device)
	for i, page in enumerate(PDFPage.get_pages(fp)):

		if page is not None:
			interpreter.process_page(page)

	device.close()
	fp.close()

	return outfp.getvalue()


def getData(filepath, sheet): 
	
	separator = ' '
	threshold = 2.5
	text = pdf_to_txt(filepath, separator, threshold).strip().split()

	# fills up incomplete date data
	sideDate = "" 
	while (text[0].lower() != "weekly"):
		sideDate += text.pop(0)
		sideDate += " "
	
	#update pointer	
	while (text[0].lower() == "weekly") or (text[0].lower() == "market" or text[0].lower() == "review"):
		text.pop(0)

	# extract date
	c = 0
	while ("201" not in text[c] and "." not in text[c]) or len(text[c]) != 4:
		c += 1
	c += 1
	date = " ".join(text[0:c])

	#update pointer
	r = 0
	for word in text:
		if word.lower() == "this":
			break
		r += 1
	r += 4

	location = True #is current text part of location data
	locations = []
	data = []
	locString = ""
	dataStr = ""
	this = True
	places = ["colombia" , "brazil", "honduras", "vietnam", "indonesia", "kenya", "ago"]
	stats = ["c", "ice", "liffe"]

	#Extract location/stats data based on keywords
	for j in range(r, len(text)):

		if text[j].lower() == "ago":
			continue

		if text[j].lower() in places:
			dataStr = ""
			locationString = ""
			location = True
	
		if (location == True) and not(text[j].lower() in stats):		
			locString += text[j] + " "
		
		if location == False:		
			dataStr += text[j] + " "      
				
		if (text[j].lower() in stats) and location == True:
			locations.append(locString)
			locString = ""
			location = False
			dataStr = text[j] + " "	
	
		if location == False and ((text[j].isdigit() or text[j][1:].isdigit()) or (text[j].lower() =="level" or text[j].lower() =="even")):	
			if this == True:
				data.append(dataStr)
				this = False
				dataStr = ""
		
			else:
				data.append(dataStr)
				this = True
				dataStr = ""
				location = True
	
	locations[0] = "brazil swedish"
			
	return([locations, data, date.decode('utf8').encode('ascii', errors='ignore'), sideDate])
	
# standardize date strings into datetime format	
def getDate(date):
	year = date.split()[-1]	
	monthDict = { "January" : 1, "JANUARY" : 1, "February": 2, "FEBRUARY" : 2, "MARCH": 3, "March" : 3, "April": 4, "APRIL":4, "MAY": 5, "May": 5, "June":6, "JUNE":6, "July":7, "JULY":7, "August":8, "AUGUST":8, "September": 9, "SEPTEMBER" :9, "October": 10, "OCTOBER":10, "November":11,"NOVEMBER":11,"December":12,"DECEMBER":12}
	
	if date.split()[3][1].strip().isdigit():
		month = date.split()[0]
		monthint = monthDict.get(month)
		dayint = int(date.split()[1])
		yearint = int(year)
		startDate = datetime.datetime(year= yearint, month= monthint, day= dayint)
		return startDate
		
	else:
		startmonth = date.split()[0]
		startmonth = monthDict.get(startmonth)
		dayint = int(date.split()[1])
		yearint = int(year)
		startDate = datetime.datetime(year= yearint, month= startmonth, day= dayint)
		return startDate

# initialize Excel
if not os.path.exists(".../Desktop/volcafe.xls"):
	book = xlwt.Workbook(encoding ="utf-8")
	sheet = book.add_sheet("sheet", cell_overwrite_ok=True)
	book.save(".../Desktop/volcafe.xls")

resultlist=[]

# parse over all Volcafe files in folder
for files in os.listdir('.../Desktop/Volcafe'):

	try:
		rb = open_workbook(".../Desktop/volcafe.xls")
		wb = copy(rb)
		s = wb.get_sheet(0)

		filepath =  ".../Desktop/Volcafe/" + files
		result = getData(filepath, s)

		##save datetime data for each file into resultlist
		date = result[2]
		if "201" not in date:
			year = result[3].split(",")[-1]		
			date += ", "
			date += year

		if date.split()[2] != "-":
		
			date = date.split()[0:2] +["-"] + date.split()[2:]
			date = " ".join(date)

		result[2] = date
		result.append(getDate(date))		
		resultlist.append(result)

		wb.save(".../Desktop/volcafe.xls")

	except:
		continue

#sort files by date
resultlist.sort(key = lambda x: x[4])

#write data into Excel
startRow = 0
for result in resultlist:
	try:
		rb = open_workbook(".../Desktop/volcafe.xls")
		wb = copy(rb)
		s = wb.get_sheet(0)

		rowl = startRow + 2
		rowd = startRow + 2
	
		location = result[0]
	
		for i in location:
			if i.strip() != "":
				s.write(rowl, 0, i)
				rowl+=1
		
		data = result[1]
	
		for j in range(len(data)):
			if j %2 == 0:
				s.write(rowd, 1, data[j])
			else:
				s.write(rowd, 3, data[j])
				rowd+=1		


		date = result[2]

		s.write(startRow,0, date)			
		s.write(startRow + 1,1, "this week")
		s.write(startRow + 1,3, "last week")
		
		startRow += 10

		wb.save(".../Desktop/volcafe.xls")

	except:
		continue

		












