#!/usr/bin/python
__author__ = 'samirw'

#################################
#	 Multi-party ZOPA Code   	#
#	 Given an excel sheet    	#
#	 with scores per option,	#
#	 returns possible packages	#
#################################

import xlrd
import xlwt
import itertools
from operator import add

#########################
#	Set Options Here	#
#########################
data_file 	= 'Hesperia.xlsx' # Excel file with scores/party/option
num_opts 	= (3, 3, 4, 4, 5) # Number of options per subpackage
min_parties = 5	# Number of parties needed to pass package
export 		= True # Export results into Excel file

# BATNAs
BATNA_list 	= (10, 10, 10, 10, 30) # Minimum score/party needed to approve package

def main():

	# Function to export results to Output File
	def export_results(results):
		# Create workbook and worksheet
		wb = xlwt.Workbook()
		ws = wb.add_sheet('ZOPA')

		# Number of subpackages
		len_sub = len(num_opts)

		# Create heading
		ws.write(0, 0, 'Package')
		for i in range(1, len_sub+1):
			ws.write(0, i, 'Package option %d' % i)
		for i in range(len_sub+2, num_parties+len_sub+2):
			ws.write(0, i, 'Player %d Score' % i)

		# Iterate through results, write in worksheet
		_row = 1
		for package, score in results.items():
			ws.write(_row, 0, str(package))
			for i in range(1, len_sub+1):
				ws.write(_row, i, package[i-1])
			for i in range(len_sub+2, num_parties+len_sub+2):
				ws.write(_row, i, score[i-(len_sub+2)])
			_row += 1

		# Save file
		wb.save('Output.xls')

	# Import excel file with scores
	wb = xlrd.open_workbook(data_file)
	score_sheet = wb.sheet_by_name('Master Scores')

	# Create necessary variables
	results 	 = {} # Results dictionary
	num_parties  = score_sheet.ncols-1 # Number of parties in the negotiation
	max_fail 	 = num_parties - min_parties # Maximum number of failures allowed in ZOPA
	score_matrix = [] # Create score matrix

	# Import data from Excel sheet
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
		scores = [0 for x in BATNA_list]
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

		#######################
		# 					  #
		#					  #
		# Custom Options Here #
		#					  #
		#					  #
		#######################

		if fails > max_fail: # Package not viable
			continue
		else: # Package passes
			package_correction = tuple([x+1 for x in package]) # Add one to every option for clarity
			results[package_correction] = scores # Add package to results dictionary

	# Export scores into Excel file
	if export:
		export_results(results)

	return results

if __name__ == "__main__":
	main()