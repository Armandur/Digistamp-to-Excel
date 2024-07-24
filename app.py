from PyPDF2 import PdfReader
from pprint import pprint
import pandas as pd
import os
from enum import Enum
from datetime import date, datetime, timedelta
import copy
from glob import glob

Digistamps = []

class Postage(Enum):
	Brev = 1
	Paket = 2
	

def swedish_month_to_number(month):
	month_dict = {
		"januari": 1,
		"februari": 2,
		"mars": 3,
		"april": 4,
		"maj": 5,
		"juni": 6,
		"juli": 7,
		"augusti": 8,
		"september": 9,
		"oktober": 10,
		"november": 11,
		"december": 12
	}
	
	return month_dict.get(month.lower(), "Invalid month")


class Digistamp():
	def __init__(self, rows = [[], [], []], postage_type=Postage.Brev, max_weight=50, post_by=(datetime.now() + timedelta(days=10)).date()):
		self.rows = rows
		self.postage_type = postage_type
		self.max_weight = max_weight
		self.post_by = post_by

	def __str__(self) -> str:
		string = '\n'.join(self.rows)
		string += f"\n{self.postage_type} - {self.max_weight} g\n"
		string += f"Posta senast {self.post_by}"
		return string


def process_input(filename):
	text = ""

	pdf = PdfReader(filename)
	for page in pdf.pages:
		text += page.extract_text()
	
	text = text.splitlines()
	#pprint(text)

	i = 0
	while i < len(text):
		#pprint(text[i])
		if "Din kod" in text[i]:
			#pprint("Found header!")
			break
		i += 1

	while i < len(text):
		stamp = Digistamp()
		if "mottagarens namn och adress" in text[i]:
			# At end of codes
			i = len(text)
			continue

		if len(text[i+1]) == 4:
			stamp.rows[0] = text[i+1]
		if len(text[i+2]) == 4:
			stamp.rows[1] = text[i+2]
		if ' g' in text[i+3]:
#			print("Found row with type of postage")
			stamp.rows[2] = text[i+3][0:4]
			type_row = text[i+3][4:].split(' - ')
			stamp.postage_type = Postage[type_row[0]]
			stamp.max_weight = int(type_row[1].split(' ')[0])

		if 'Giltig' in text[i+4]:
			post_by = text[i+4].split(" senast ")[1].split(" ")
			year = 0
			month = 0
			day = 0
			if len(post_by[2]) == 4:
				year = int(post_by[2])
			else:
				year = int(post_by[2][0:4])
				if len(post_by[2]) == 8:
					nextrow = post_by[2][4:]
					text.insert(i+5,nextrow)
				elif len(post_by[2]) > 8:
					#Have reached end of codes
					i = len(text)

			
			month = swedish_month_to_number(post_by[1])
			day = int(post_by[0])
			
			stamp.post_by = date(year, month, day)
			
			#pprint(stamp.rows)
			#pprint(stamp.postage_type)
			#pprint(stamp.max_weight)
			#pprint(stamp.post_by)
			
			Digistamps.append(copy.deepcopy(stamp))
			
			i += 4
			continue
		i += 1

	#pprint(text)
	return Digistamps

def write_output(digistamps):
	file = "output.xlsx"
	columns = ["type", "max_weight", "post_by", "row1", "row2", "row3"]
	df = pd.DataFrame(columns=columns)
	
	def append_row_to_dataframe(df, row_data):
		new_row = pd.DataFrame([row_data], columns=columns)
		return pd.concat([df, new_row], ignore_index=True)
	
	for stamp in digistamps:
		df = append_row_to_dataframe(df, [stamp.postage_type, stamp.max_weight, stamp.post_by, stamp.rows[0], stamp.rows[1], stamp.rows[2]])

	pprint(df)
	df.to_excel(file, index=False)
	
pdfs = glob(os.path.join(os.getcwd(), "*.pdf"))
for pdf in pdfs:
	process_input(pdf)

#for stamp in Digistamps:
	#print(stamp)

write_output(Digistamps)

pprint(f"Processed {len(Digistamps)} digistamsp from {len(pdfs)} pdfs.")

# Find first row with Din kod
# Start iterating and begin collecting rows for Digistamp
# If len(row) == 4, add row to Digistamp
# If row ends with " g" remove all but first four letters, add to Digistamp, Split rest by ' - ', first element is type of postage, second is max weight.
# If row begins with "Giltig", remove '\n', split by ' ', get last three elements and convert to date (2 augist 2024), add to Digistamp
