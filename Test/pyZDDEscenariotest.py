#-------------------------------------------------------------------------------
# Name:        pyZDDEscenariotest.py
# Purpose:     To test different scenerio such as instantiating multiple ZEMAX
#              connections and other scenerios. This test doesnot use the python
#              unit testing framework
#
# Author:      Indranil Sinharoy
#
# Created:     19/10/2012
# Copyright:   (c) XPMUser 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os, sys, time

# Put both the "Test" and the "PyZDDE" directory in the python
# search path.
testdirectory = os.getcwd()
ind = testdirectory.find('Test')
pyzddedirectory = testdirectory[0:ind-1]
if testdirectory not in sys.path:
    sys.path.append(testdirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

# Import the pyZDDE module
import pyZDDE

def testSetup():
    # Setup up the basic environment for the scenerio test
    #ToDo: automatically open ZEMAX
    pass

def test_scenario_multipleChannel():
    """Test multiple channels of communication with ZEMAX"""
    # Create multiple client objects
    link0 = pyZDDE.pyzdde()
    link1 = pyZDDE.pyzdde()
    link2 = pyZDDE.pyzdde()

    # Initialize
    ch0_status = link0.zDDEInit()
    print "Status for link 0:", ch0_status
    assert ch0_status==0
    time.sleep(0.25) # Not required, but just for observation

    ch1_status = link1.zDDEInit()
    print "Status for link 1:",ch1_status
    #assert ch1_status == 0   #FIXME: Unable to create second communication link.
    time.sleep(0.25)   # Not required, but just for observation

    # Call somefunctions to do something
    if not ch0_status:
        #if channel 0 successfully created
        link0.zNewLens()
        sysPara = link0.zGetSystem()
        sysParaNew = link0.zSetSystem(0,sysPara[2],0,0,20,1,-1)
        link0.zInsertSurface(1)

    # Delete one of the objects
    del link2





    # Close the channels
    link0.zDDEClose()
    link1.zDDEClose()

if __name__ == '__main__':
    # Set the python optimization flag (generally this should be "False"
    # for the testing
    testSetup()
    test_scenario_multipleChannel()
