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


	"""
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
			print(casename)
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
			

			auxtext = open('checkWv3.aux', 'r').read().replace('CASENAME', casename)
			
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
"""
	casename = sys.argv[1]

	print(casename)
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
			

	auxtext = open('checkWv3.aux', 'r').read().replace('CASENAME', casename)
			
	print('Start an powerworld case.....')
	#raw_name = casename + '-W.raw'
	raw_object = pypowerworld.PyPowerWorld(casename + '-W.pwb')
	#raw_object.opencase(raw_name)
	print('Start loading auxtext.....')

	raw_object.loadauxfiletext(auxtext)

	print('Finish loading auxtext')
	raw_object.closecase()
	print('Successfully ran aux text.')
""""""



class PyPowerWorld(object):
    """Class object designed for easy interface with PowerWorld."""
    def __init__(self, pwb_file_path=None):
        try:
            self.__pwcom__ = win32com.client.Dispatch('pwrworld.SimulatorAuto')
        except Exception as e:
            print(str(e))
            print("Unable to launch SimAuto.",
                  "Please confirm that your PowerWorld license includes the SimAuto add-on ",
                  "and that SimAuto has been successfuly installed.")
        self.pwb_file_path = pwb_file_path
        self.__setfilenames__()
        self.output = ''
        self.error = False
        self.error_message = ''
        self.COMout = ''

    def __setfilenames__(self):
        self.file_folder = os.path.split(self.pwb_file_path)[0]
        self.file_name = os.path.splitext(os.path.split(self.pwb_file_path)[1])[0]
        self.aux_file_path = self.file_folder + '/' + self.file_name + '.aux'  # some operations require an aux file
        self.save_file_path = os.path.splitext(os.path.split(self.pwb_file_path)[1])[0]

    def __pwerr__(self):
        if self.COMout is None:
            self.output = None
            self.error = False
            self.error_message = ''
        elif self.COMout[0] == '':
            self.output = None
            self.error = False
            self.error_message = ''
        elif 'No data' in self.COMout[0]:
            self.output = None
            self.error = False
            self.error_message = self.COMout[0]
        else:
            self.output = self.COMout[-1]
            self.error = True
            self.error_message = self.COMout[0]
        return self.error            
    
    def opencase(self, pwb_file_path):
        """Opens case defined by the full file path; if this is undefined, opens by previous file path"""
        self.COMout = self.__pwcom__.OpenCase(pwb_file_path)
        if self.__pwerr__():
        	print('Error opening case:\n\n%s\n\n', self.error_message)
        	print('Please check the file name and path and try again (using the opencase method)\n')
        	return True
    
            
    def closecase(self):
        """Closes case without saving changes."""
        self.COMout = self.__pwcom__.CloseCase()
        if self.__pwerr__():
            print('Error closing case:\n\n%s\n\n', self.error_message)
            return False
        return True

    
    def loadauxfiletext(self,auxtext):
        """Creates and loads an Auxiliary file with the text specified in auxtext parameter."""
        f = open(self.aux_file_path, 'w')
        f.writelines(auxtext)
        f.close()
        self.COMout = self.__pwcom__.ProcessAuxFile(self.aux_file_path)
        if self.__pwerr__():
            print('Error running auxiliary text:\n\n%s\n', self.error_message)
            return False
        return True
    
    def exit(self):
        """Clean up for the PowerWorld COM object"""
        self.closecase()
        del self.__pwcom__
        self.__pwcom__ = None
        return None
    
    def __del__(self):
        self.exit()
		



if __name__ == '__main__':
	run_contingency_analysis()