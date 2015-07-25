# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        pyZDDEscenariotest.py
# Purpose:     To test different scenerio such as instantiating multiple ZEMAX
#              connections and other scenerios. This test doesnot use the python
#              unit testing framework
#
# Licence:     MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function
import os
import sys
import time

# Put both the "Test" and the "PyZDDE" directory in the python search path.
testdirectory = os.path.dirname(os.path.realpath(__file__))
ind = testdirectory.find('Test')
pyzddedirectory = testdirectory[0:ind-1]
if testdirectory not in sys.path:
    sys.path.append(testdirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

# Import the pyzdde module
#import pyzdde
import pyzdde.zdde as pyz

# ZEMAX file directory
zmxfp = pyzddedirectory+'\\ZMXFILES\\'

def testSetup():
    # Setup up the basic environment for the scenerio test
    #ToDo: automatically open ZEMAX
    pass

def test_scenario_multipleChannel():
    """Test multiple channels of communication with ZEMAX"""
    # Create multiple client objects
    ln0 = pyz.PyZDDE()
    ln1 = pyz.PyZDDE()
    ln2 = pyz.PyZDDE()

    # Initialize
    ch0_status = ln0.zDDEInit()
    print("\nStatus for link 0:", ch0_status)
    assert ch0_status==0  # Note that this will case the program to terminate without shutting down the server. The program will have to be restarted.
    print("App Name for Link 0:", ln0._appName)
    print("Connection status for Link 0:", ln0._connection)
    time.sleep(0.1) # Not required, but just for observation

    ch1_status = ln1.zDDEInit()
    print("\nStatus for link 1:",ch1_status)
    assert ch1_status == 0 # Note that this will case the program to terminate without shutting down the server. The program will have to be restarted.
    print("App Name for Link 1:", ln1._appName)
    print("Connection status for Link 1:", ln1._connection)
    time.sleep(0.1)   # Not required, but just for observation

    # Create a new lens in the first ZEMAX DDE server
    ln0.zNewLens()
    sysPara = ln0.zGetSystem()
    sysParaNew = ln0.zSetSystem(0,sysPara[2], 0, 0, 20, 1, -1)
    ln0.zInsertSurface(1)

    # Delete one of the objects
    del ln2   # We can delete this object like this since no DDE conversation object was created for it.

    # Load a lens into the second ZEMAX DDE server
    filename = zmxfp+"Cooke 40 degree field.zmx"
    ret = ln1.zLoadFile(filename)
    assert ret == 0
    print("\nzLoadFile test successful")


    #Get system from both the channels
    recSys0Data = ln0.zGetSystem()
    recSys1Data = ln1.zGetSystem()

    print("\nSystem data from 1st system:", recSys0Data)
    print("\nSystem data from 2nd system:", recSys1Data)


    # Close the channels
    ln0.zDDEClose()
    ln1.zDDEClose()

if __name__ == '__main__':
    # Set the python optimization flag (generally this should be "False"
    # for the testing
    testSetup()
    test_scenario_multipleChannel()
