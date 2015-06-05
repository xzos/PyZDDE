# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        checkDataItemCompleteness.py
# Purpose:     Helper code to check if the current version of PyZDDE contains
#              all the data items defined by ZEMAX
#
#              NOTE:
#              A list of data items are available in the Extensions chapter of
#              the ZEMAX manual or the OpticStudio Help
#
# Licence:     MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function
import os
import inspect
import pyzdde.zdde as pyz

testdirectory = os.path.dirname(os.path.realpath(__file__))

def main():
    # Get data items (class methods) from PyZDDE
    pyzobj = pyz.PyZDDE()
    method_list = [method[0] for method in inspect.getmembers(pyzobj, predicate=inspect.ismethod)]
    dataItemSet_pyzdde = []
    ipzFunctions = [] # interactive helper functions
    for item in method_list:
        if item.startswith('z'):
            dataItemSet_pyzdde.append(item.split('z', 1)[1])
        elif item.startswith('ipz'):
            ipzFunctions.append(item)
    dataItemSet_pyzdde = set(dataItemSet_pyzdde)
    #print("\nData items in PyZDDE: \n", dataItemSet_pyzdde)
    itemCount_pyzdde = len(dataItemSet_pyzdde)

    # Get data items from textfile
    dataItemSet_zemax = []
    dataItemFile = open(testdirectory+os.path.sep+"zemaxDataItems.txt","r")
    for line in dataItemFile:
        if line.rstrip() is not '':
            if not line.rstrip().startswith('#'):
                dataItemSet_zemax.append(line.rstrip())
    dataItemFile.close()
    dataItemSet_zemax = set(dataItemSet_zemax)

    # Obsolete data items
    dataItemSet_obsolete = set(['SetUDOData'])
    dataItemSet_zemax = dataItemSet_zemax - dataItemSet_obsolete
    #print("\nData items in Zemax: \n", dataItemSet_zemax)
    itemCount_zemax = len(dataItemSet_zemax)

    # Print useful information
    print("Total number of data items defined in ZEMAX Manual:", itemCount_zemax)
    print("Total number of PyZDDE class methods with prefix z:", itemCount_pyzdde)

    if dataItemSet_zemax.issubset(dataItemSet_pyzdde):
        print("\nGREAT! All data items defined in Zemax Manual are in PyZDDE.")
    else:
        dataItemsNotInPyZDDE = dataItemSet_zemax.difference(dataItemSet_pyzdde)
        print("\nOOPS! Looks like there are some still lurking around.")
        print("\nData items in Zemax but not in PyZDDE: \n", dataItemsNotInPyZDDE)

    # Other functions not defined in ZEMAX manual
    extraFunctions_set = dataItemSet_pyzdde.difference(dataItemSet_zemax)
    extraFunctions = [''.join(['z', exfun]) for exfun in extraFunctions_set]
    print("\Extra functions:")
    print(extraFunctions)
    print("Total extra functions: ", len(extraFunctions))
    print("\nipz helper functions:")
    print(ipzFunctions)
    print("Total ipz functions: ", len(ipzFunctions))

if __name__ == '__main__':
    main()
