#-------------------------------------------------------------------------------
# Name:      beamFileRead.py
# Purpose:   Demonstrate the function readBeamFile() that reads a zemax beam
#            file.
#
# NOTE:      Please note that this code uses matplotlib plotting library from
#            http://matplotlib.org/ for 2D-plotting
#
# Copyright: (c) 2012 - 2014
# Licence:   MIT License
#-------------------------------------------------------------------------------
from pyzdde.zdde import readBeamFile
import os
import matplotlib.pyplot as plt

directory = os.path.dirname(os.path.realpath(__file__))

beamfilename = directory+os.sep+"type2spdc.zbf"

beamData = readBeamFile(beamfilename)

(version, (nx, ny), ispol, units, (dx,dy), (zposition_x, zposition_y),
 (rayleigh_x, rayleigh_y), (waist_x, waist_y), lamda, index,
 (x_matrix, y_matrix), (Ex_real, Ex_imag, Ey_real, Ey_imag)) = beamData

xlabels = [-nx*dx/2+x*dx for x in range(0, nx)]
ylabels = [-ny*dy/2+y*dy for y in range(0, ny)]

fig, ax = plt.subplots()

cs1 = ax.contourf(xlabels, ylabels, Ey_real, cmap="Blues", alpha=0.66)
cs2 = ax.contourf(xlabels, ylabels, Ex_real, cmap="Reds", alpha=0.66)

ax.set_xlabel("X position (mm)")
ax.set_ylabel("Y position (mm)")
plt.show()
