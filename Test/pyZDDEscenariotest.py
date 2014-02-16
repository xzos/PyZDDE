#-------------------------------------------------------------------------------
# Name:        pyZDDEscenariotest.py
# Purpose:     To test different scenerio such as instantiating multiple ZEMAX
#              connections and other scenerios. This test doesnot use the python
#              unit testing framework
#
# Copyright:   (c) Indranil Sinharoy, 2012 - 2014
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
import pyzdde.zdde

# ZEMAX file directory
zmxfp = pyzddedirectory+'\\ZMXFILES\\'

def testSetup():
    # Setup up the basic environment for the scenerio test
    #ToDo: automatically open ZEMAX
    pass

def test_scenario_multipleChannel():
    """Test multiple channels of communication with ZEMAX"""
    # Create multiple client objects
    link0 = pyzdde.zdde.PyZDDE()
    link1 = pyzdde.zdde.PyZDDE()
    link2 = pyzdde.zdde.PyZDDE()

    # Initialize
    ch0_status = link0.zDDEInit()
    print("Status for link 0:", ch0_status)
    assert ch0_status==0  # Note that this will case the program to terminate without shutting down the server. The program will have to be restarted.
    print("App Name for Link 0:", link0.appName)
    print("Connection status for Link 0:", link0.connection)
    time.sleep(0.1) # Not required, but just for observation

    ch1_status = link1.zDDEInit()
    print("Status for link 1:",ch1_status)
    assert ch1_status == 0 # Note that this will case the program to terminate without shutting down the server. The program will have to be restarted.
    print("App Name for Link 1:", link1.appName)
    print("Connection status for Link 1:", link1.connection)
    time.sleep(0.1)   # Not required, but just for observation

    # Create a new lens in the first ZEMAX DDE server
    link0.zNewLens()
    sysPara = link0.zGetSystem()
    sysParaNew = link0.zSetSystem(0,sysPara[2],0,0,20,1,-1)
    link0.zInsertSurface(1)

    # Delete one of the objects
    del link2   # We can delete this object like this since no DDE conversation object was created for it.

    # Load a lens into the second ZEMAX DDE server
    filename = zmxfp+"Cooke 40 degree field.zmx"
    ret = link1.zLoadFile(filename)
    assert ret == 0
    print("zLoadFile test successful")


    #Get system from both the channels
    recSys0Data = link0.zGetSystem()
    recSys1Data = link1.zGetSystem()

    print("System data from 1st system:\n",recSys0Data)
    print("System data from 2nd system:\n",recSys1Data)


    # Close the channels
    link0.zDDEClose()
    link1.zDDEClose()

if __name__ == '__main__':
    # Set the python optimization flag (generally this should be "False"
    # for the testing
    testSetup()
    test_scenario_multipleChannel()
