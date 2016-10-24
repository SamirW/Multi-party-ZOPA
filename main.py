#!/usr/bin/python
__author__ = 'samirw'

#################################
#	 Multi-party ZOPA Code   	#
#	 Given an excel sheet    	#
#	 with scores per option		#
#	 returns possible packages	#
#################################

import xlrd

# Options
min_parties = 5	# Number of minimum parties to pass package
data_file 	= 'Hesperia.xlsx' # Excel file with scores/party/option

# BATNAs
BATNA_list = (10, 10, 10, 10, 10, 10) # Minimum score/party needed to approve package

def main(self):
