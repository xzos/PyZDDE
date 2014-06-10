#-------------------------------------------------------------------------------
# Name:      couplingEfficiencySingleModeFibers.py
# Purpose:   Demonstrate the zGetPOP() function of pyZDDE.
#            Calculates the coupling efficiency between two single mode fibers 
#            at 810 nm.
#
# NOTE:      Please note that this code uses matplotlib plotting library from
#            http://matplotlib.org/ for 2D-plotting
#
# Copyright: (c) 2012 - 2014
# Licence:   MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function
import pyzdde.zdde as pyz
import matplotlib.pyplot as plt
import os

ln = pyz.createLink()

directory = os.path.dirname(os.path.realpath(__file__))
popCfgFile = directory + "\POP.CFG"
popOutputFile = directory + "\popInfo.txt" 

# let's define our gaussian beam first
# make the sampling grid 128 by 128
ln.zModifySettings(popCfgFile, "POP_SAMPX", 3)        
ln.zModifySettings(popCfgFile, "POP_SAMPY", 3)

# make our source beam have a waist of 3.25 microns, with a divergence of 
# 3.25 degrees, mimicking a single mode fiber 

ln.zModifySettings(popCfgFile, "POP_BEAMTYPE", 2)
ln.zModifySettings(popCfgFile, "POP_PARAM0", 0.00175)
ln.zModifySettings(popCfgFile, "POP_PARAM1", 0.00175)
ln.zModifySettings(popCfgFile, "POP_PARAM3", 3.25)
ln.zModifySettings(popCfgFile, "POP_PARAM4", 3.25)


# do the same for the target fiber

ln.zModifySettings(popCfgFile, "POP_FIBERTYPE", 2)
ln.zModifySettings(popCfgFile, "POP_FPARAM0", 0.00175)
ln.zModifySettings(popCfgFile, "POP_FPARAM1", 0.00175)
ln.zModifySettings(popCfgFile, "POP_FPARAM3", 3.25)
ln.zModifySettings(popCfgFile, "POP_FPARAM4", 3.25)
ln.zModifySettings(popCfgFile, "POP_COMPUTE", 1)

# run the POP
[peakIrradiance, totalPower,fiberEfficiency_system,fiberEfficiency_receiver,coupling,pilotSize,pilotWaist,pos,rayleigh,powerGrid] = ln.zGetPOP(popOutputFile,True,popCfgFile)
ln.close()

print("Fiber coupling efficiency for this system is: ", coupling)

# plot the beam
plt.imshow(powerGrid)
plt.show()

