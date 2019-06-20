import csv
import pdb
import sys
import os
import pypowerworld

import pandas as pd
import numpy as np
import win32com
from win32com.client import VARIANT
import pythoncom
#Test.inl is a sample inl file
#Creates TestINL.csv with only first, second, and sixth columns


def run_contingency_analysis():

	files = []
	casenames = []
	for r, d, f in os.walk(sys.argv[1]):
		for file in f:
			if '.PWB' in file:
				files.append(os.path.join(r, file))

	for file in files:
		casenames.append(file)

	for casename in casenames:
		print(casename)

	try:
		for casename in casenames:
			print(casename)
			
			auxtext = open('check_pwb_ctg.aux', 'r').read().replace('CASENAME', casename)
			
			print('Start an powerworld case.....')
			#raw_name = casename + '-W.raw'
			raw_object = pypowerworld.PyPowerWorld(casename)
			#raw_object.opencase(raw_name)
			print('Start loading auxtext.....')

			raw_object.loadauxfiletext(auxtext)

			print('Finish loading auxtext')
			raw_object.closecase()
			print('Successfully ran aux text.')

	except Exception as e:
            print(e) 

		



if __name__ == '__main__':
	run_contingency_analysis()