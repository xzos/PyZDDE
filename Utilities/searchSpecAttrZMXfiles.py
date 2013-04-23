#-------------------------------------------------------------------------------
# Name:        searchSpecAttrZMXfiles.py
# Purpose:     Search for a specific attribute through ZMX files
#
# Author:      Indranil
#
# Created:     23/04/2013
# Copyright:   (c) Indranil 2013
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os, glob, sys

# Put both the "Utilities" and the "PyZDDE" directory in the python search
# path.
utilsDirectory = os.getcwd()
ind = utilsDirectory.find('Utilities')
pyzddedirectory = utilsDirectory[0:ind-1]
if utilsDirectory not in sys.path:
    sys.path.append(utilsDirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

import pyZDDE

# ZEMAX file directory to search
zmxfp = "C:\\PROGRAMSANDEXPERIMENTS\\ZEMAX\\Samples\\Sequential\\Objectives\\"
pattern = "*.zmx"

# Create a DDE channel object
pyZmLnk = pyZDDE.pyzdde()
#Initialize the DDE link
stat = pyZmLnk.zDDEInit()

filenames = glob.glob(zmxfp+pattern)
print filenames[1:2]


#Close the DDE channel
pyZmLnk.zDDEClose()










def main():
    pass

if __name__ == '__main__':
    main()
