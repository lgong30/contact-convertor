#!/usr/bin/python
# -*- coding: utf-8 -*-

"""Convert Gmail (Outlook csv format) contact to Outlook csv format contact

This script converts the csv file exported from gmail into the csv format
can be recognized by Outlook (web version) Importer, and append them to 
Outlook Contact.

To use it, you need first export your contact from Gmail.
"""

import csv


# Gmail contact 
gmc = 'contacts-gmc.csv'
# Outlook contact header
olc = 'contacts-olc.csv'


olc_header = csv.reader(open(olc, 'r')).next()
olc_header[0] = 'First Name'

# create a map from each header entry to its index
n = len(olc_header)
mapper_olc = dict(zip(olc_header, range(n)))

# create a csv writer
fp = open(olc, 'a')
writer = csv.writer(fp, delimiter=',', lineterminator='\n')

def get_X_index(header, X):
	"""Get the index of X
	"""
	index = []
	for i, elm in enumerate(header):
		if elm.lower().find(X) != -1:
			index.append(i)
	return index

def has_X(row, X_index):
	"""Check whether a row has X 
	"""
	for index in X_index:
		if not row[index]:
			pass
		else:
			return True
	return False



# process gmail csv file
with open(gmc, 'r') as csvf:
	reader = csv.reader(csvf)
	header = reader.next()

	names_index = get_X_index(header, 'name')
	email_index = get_X_index(header, 'e-mail')
	phone_index = get_X_index(header, 'phone')
	
	for row in reader:
		# only consider rows with at least one of First Name,
		# Last Name, and Middle Name, and at least with an e-mail
		# or a phone number
		if has_X(row, names_index) and \
		   (has_X(row, email_index) or has_X(row, phone_index)):
			olc_row = [''] * n
			for i, elm in enumerate(row):
				if not elm:
					pass
				else:
					if mapper_olc.has_key(header[i]):
						col_id = mapper_olc[header[i]]
						olc_row[col_id] = elm
			writer.writerow(olc_row)
# close the file
fp.close()



