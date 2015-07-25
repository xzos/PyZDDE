#-------------------------------------------------------------------------------
# Name:      couplingEfficiencySingleModeFibers.py
# Purpose:   Demonstrate the following function related to POP in Zemax:
#            zGetPOP(), zSetPOPSettings(), zModifyPOPSettings()
#            Calculates the fiber coupling efficiency between a Gaussian beam
#            in free space that is focused into a fiber.
#
# NOTE:      Please note that this code uses matplotlib plotting library from
#            http://matplotlib.org/ for 2D-plotting
#
# Copyright: (c) 2012 - 2014
# Licence:   MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import pyzdde.zdde as pyz
import matplotlib.pyplot as plt
import math
import os

ln = pyz.createLink()

curDir = os.path.dirname(os.path.realpath(__file__))
samplesDir = ln.zGetPath()[1]
lens = ['Simple Lens.zmx', 'Fiber Coupling.zmx']
popfile = os.path.join(samplesDir, 'Physical Optics', lens[0])
cfgFile = os.path.join(curDir, "coupEffSgleModePOPEx.CFG")

# load pop file into Zemax server
ln.zLoadFile(popfile)

# source gaussian beam (these parameters can be deduced using the paraxial
# gaussian beam calculator under the Analyze menu in Zemax):
#    beam type = Gaussian Size + Angle;
#    size-x/y (beam waist) = 2 mm;
#    angle x/y in degrees (divergence) = 0.00911890
#    Tot power = 1
# fiber coupling integral parameters:
#    beam type = Gaussian Size + Angle;
#    size-x/y = 0.008 mm;
#    angle x/y in degrees (divergence) = 2.290622
# display parameters
# sampling grid 256 by 256; x/y-width = 40 by 40;

srcParam = ((1, 2, 3, 4), (2, 2, 0.00911890, 0.00911890))
fibParam = ((1, 2, 3, 4), (0.008, 0.008, 2.290622, 2.290622))

# Setup POP analysis
ln.zSetPOPSettings(data=0, settingsFile=cfgFile, startSurf=1, endSurf=1,
                   field=1, wave=1, beamType=2, paramN=srcParam, tPow=1,
                   sampx=4, sampy=4, widex=40, widey=40, fibComp=1, fibType=2,
                   fparamN=fibParam)

# compute and get POP data (irradiance) at the source surface
popInfo_src_irr, data_src_irr = ln.zGetPOP(settingsFile=cfgFile, displayData=True)

# modify the POP settings to display the irradiance plot at the focused point
# (end_surf = 4)
errStat = ln.zModifyPOPSettings(cfgFile, endSurf=4)
print('Modify Settings: errStat =', errStat)

# get data at fiber
popInfo_dst_irr, data_dst_irr =  ln.zGetPOP(settingsFile=cfgFile, displayData=True)

# modify the POP settings to get Phase data at the source surface. Note that
# when changing the data type, we need to pass all settings again.
ln.zSetPOPSettings(data=1, settingsFile=cfgFile, startSurf=1, endSurf=1,
                   field=1, wave=1, beamType=2, paramN=srcParam, tPow=1,
                   sampx=4, sampy=4, widex=40, widey=40, fibComp=1, fibType=2,
                   fparamN=fibParam)

# compute and get the POP phase data
popInfo_src_phase, data_src_phase =  ln.zGetPOP(settingsFile=cfgFile, displayData=True)
# again, modify the POP settings to display the phase plot at the focused point
# (end_surf = 4)
errStat = ln.zModifyPOPSettings(cfgFile, endSurf=4)
print('Modify Settings: errStat =', errStat)
# get phase data at fiber
popInfo_dst_phase, data_dst_phase =  ln.zGetPOP(settingsFile=cfgFile, displayData=True)

# close the DDE link
ln.close()

# print useful information
print("\nPop information (irradiance) at the source surface: ")
print(popInfo_src_irr)
print("\nPop information (irradiance) at the fiber surface: ")
print(popInfo_dst_irr)
print("\nCoupling efficiency: ", popInfo_dst_irr[4])

# plot the beam at the source and at the fiber
fig = plt.figure(facecolor='w')

# irradiance data
ax = fig.add_subplot(2,2,1)
ax.set_title('Irradiance at source')
ext = [-popInfo_src_irr.widthX/2, popInfo_src_irr.widthX/2,
       -popInfo_src_irr.widthY/2, popInfo_src_irr.widthY/2]
ax.imshow(data_src_irr, extent=ext, origin='lower')
ax.set_xlabel('x (mm)'); ax.set_ylabel('y (mm)')

ax = fig.add_subplot(2,2,2)
ax.set_title('Irradiance at fiber')
ext = [-popInfo_dst_irr.widthX/2, popInfo_dst_irr.widthX/2,
       -popInfo_dst_irr.widthY/2, popInfo_dst_irr.widthY/2]
ax.imshow(data_dst_irr, extent=ext, origin='lower')
ax.set_xlabel('x (mm)'); ax.set_ylabel('y (mm)')

# phase data
ax = fig.add_subplot(2,2,3)
ax.set_title('Phase at source')
ext = [-popInfo_src_phase.widthX/2, popInfo_src_phase.widthX/2,
       -popInfo_src_phase.widthY/2, popInfo_src_phase.widthY/2]
ax.imshow(data_src_phase, extent=ext, origin='lower',
          vmin=-math.pi, vmax=math.pi)
ax.set_xlabel('x (mm)'); ax.set_ylabel('y (mm)')

ax = fig.add_subplot(2,2,4)
ax.set_title('Phase at fiber')
ext = [-popInfo_dst_phase.widthX/2, popInfo_dst_phase.widthX/2,
       -popInfo_dst_phase.widthY/2, popInfo_dst_phase.widthY/2]
ax.imshow(data_dst_phase, extent=ext, origin='lower',
          vmin=-math.pi, vmax=math.pi)
ax.set_xlabel('x (mm)'); ax.set_ylabel('y (mm)')

fig.tight_layout()
plt.show()