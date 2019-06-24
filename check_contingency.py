import csv
import pdb
import sys
import os
import pypowerworld
#Test.inl is a sample inl file
#Creates TestINL.csv with only first, second, and sixth columns


def run_multiple_contingency_analysis_on_raw():
	
	files = []
	casenames = []
	for r, d, f in os.walk(sys.argv[1]):
		for file in f:
			if 'V0.con' in file:
				files.append(os.path.join(r, file))

	for file in files:
		casenames.append(file.replace('.con', ''))

	for casename in casenames:
		print(casename)

	try:
		for casename in casenames:
			print('Check contingency for ' + casename)
			with open(casename + ".inl","r") as source:
				rdr= csv.reader( source )
				with open(casename + "_INL.csv","w", newline = '') as result:
					wtr= csv.writer( result )
					wtr.writerow(("Gen", "", ""))
					wtr.writerow(("BusNum","ID", "PartFact"))

					for r in rdr:
						if r[0] == '0':
							break

						wtr.writerow( (r[0], r[1], r[5]) )
			
			cwd = os.getcwd()

			auxtext = open('checkWv3.aux', 'r').read().replace('CASENAME', casename).replace('CURRENT_PATH', cwd)
			
			print('Start an powerworld case.....')
			#raw_name = casename + '-W.raw'
			raw_object = pypowerworld.PyPowerWorld(casename + '-W.pwb')
			#raw_object.opencase(raw_name)
			print('Start loading auxtext.....')

			raw_object.loadauxfiletext(auxtext)

			print('Finish loading auxtext')
			raw_object.closecase()
			print('Successfully ran aux text.')

	except Exception as e:
		print(e) 




def run_single_contingency_analysis_on_raw():

	casename = sys.argv[1]

	print('Check contingency for ' + casename)
	with open(casename + ".inl","r") as source:
		rdr= csv.reader( source )
		with open(casename + "_INL.csv","w", newline = '') as result:
			wtr= csv.writer( result )
			wtr.writerow(("Gen", "", ""))
			wtr.writerow(("BusNum","ID", "PartFact"))
			for r in rdr:
				if r[0] == '0':
					break

				wtr.writerow( (r[0], r[1], r[5]) )
			
	cwd = os.getcwd()
	auxtext = open('checkWv3.aux', 'r').read().replace('CASENAME', casename).replace('CURRENT_PATH', cwd)
			
	try:
		print('Start an powerworld case.....')
		#raw_name = casename + '-W.raw'
		raw_object = pypowerworld.PyPowerWorld(casename + '-W.pwb')
		#raw_object.opencase(raw_name)
		print('Start loading auxtext.....')

		raw_object.loadauxfiletext(auxtext)

		print('Finish loading auxtext')
		raw_object.closecase()
		print('Successfully ran aux text.')
	except Exception as e:
		print(e)





if __name__ == '__main__':
	run_single_contingency_analysis_on_raw()