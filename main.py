#!/usr/bin/python
__author__ = 'samirw'

#################################
#	 Multi-party ZOPA Code   	#
#	 Given an excel sheet    	#
#	 with scores per option		#
#	 returns possible packages	#
#################################

import xlrd
import itertools
from operator import add

# Options
num_parties = 6 # Number of parties in the negotiation
min_parties = 5	# Number of minimum parties to pass package
data_file 	= 'Hesperia.xlsx' # Excel file with scores/party/option

# BATNAs
BATNA_list 	= (10, 10, 10, 10, 10, 10) # Minimum score/party needed to approve package

# Number of options per subpackage
num_opts 	= (3, 3, 4, 4, 5) 

def main():

	# Variable to cap number of failures allowed in ZOPA
	max_fail 	= num_parties - min_parties 

	# Results dictionary
	results = {}

	# Import excel file with scores
	xl_workbook = xlrd.open_workbook(data_file)
	score_sheet = xl_workbook.sheet_by_name('Master Scores')

	# Create score matrix
	score_matrix = []

	for row in range(1, score_sheet.nrows):
		_row = []
		for col in range(1, score_sheet.ncols):
				_row.append(score_sheet.cell_value(row,col))
		score_matrix.append(_row)

	# Create all possible packages
	num_opts_lists = []	# Create list of possible subpackage options
	for num in num_opts:
		num_opts_lists.append(range(num))

	packages = list(itertools.product(*num_opts_lists))

	# Loop through packages and determine viable packages
	for package in packages:

		# Find scores per party for package
		scores = [0, 0, 0, 0, 0]
		_row = 0
		for i in range(len(num_opts)):
			option_row 	 = _row + package[i]
			option_score = score_matrix[option_row]
			scores = map(add, scores, option_score)
			_row += num_opts[i]

		# Check if score allows package to be passed
		fails = 0
		for i in range(len(scores)):
			if scores[i] < BATNA_list[i]:
				fails += 1

		if fails > max_fail:
			continue
		else:
			results[package] = scores

	return results

results = main()