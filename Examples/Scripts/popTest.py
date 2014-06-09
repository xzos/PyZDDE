#POP example
#Calculates the coupling efficiency between two single mode fibers at 810 nm.

import pyzdde.zdde as pyz
import matplotlib.pyplot as plt
import os
link = pyz.createLink()

directory = os.path.dirname(os.path.realpath(__file__))
POPConfigFilename = directory+"\POP.CFG"
POPOutputFilename = directory+"\popInfo.txt" 
#let's define our gaussian beam first
#make the sampling grid 128 by 128
link.zModifySettings(POPConfigFilename, "POP_SAMPX", 3)        
link.zModifySettings(POPConfigFilename, "POP_SAMPY", 3)
#make our source beam have a waist of 3.25 microns, with a divergence of 3.25 degrees, 
#mimicking a single mode fiber 

link.zModifySettings(POPConfigFilename, "POP_BEAMTYPE", 2)
link.zModifySettings(POPConfigFilename, "POP_PARAM0", 0.00175)
link.zModifySettings(POPConfigFilename, "POP_PARAM1", 0.00175)
link.zModifySettings(POPConfigFilename, "POP_PARAM3", 3.25)
link.zModifySettings(POPConfigFilename, "POP_PARAM4", 3.25)

#do the same for the target fiber

link.zModifySettings(POPConfigFilename, "POP_FIBERTYPE", 2)
link.zModifySettings(POPConfigFilename, "POP_FPARAM0", 0.00175)
link.zModifySettings(POPConfigFilename, "POP_FPARAM1", 0.00175)
link.zModifySettings(POPConfigFilename, "POP_FPARAM3", 3.25)
link.zModifySettings(POPConfigFilename, "POP_FPARAM4", 3.25)

#run the POP
[peakIrradiance, totalPower,fiberEfficiency_system,fiberEfficiency_receiver,coupling,pilotSize,pilotWaist,pos,rayleigh,powerGrid] = link.zGetPOP(POPOutputFilename,True,POPConfigFilename)
link.close()
#plot the beam
plt.imshow(powerGrid)
plt.show()
print("Fiber coupling efficiency for this system is: ",coupling)
