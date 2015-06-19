#-------------------------------------------------------------------------------
# Name:      plotSpiralSpot.py
# Purpose:   Example of using the "spiral spot" covenience function of pyZDDE.
#
# NOTE:      Please note that this code uses matplotlib plotting library from
#            http://matplotlib.org/ for 2D-plotting
#
# Copyright: (c) 2012- 2014
# Licence:   MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function
import os
import matplotlib.pyplot as plt
import pyzdde.zdde as pyz

ln = pyz.createLink()
filename = os.path.join(ln.zGetPath()[1], 'Sequential', 'Objectives', 
                        'Cooke 40 degree field.zmx')
# Load a lens file into the ZEMAX DDE server
ln.zLoadFile(filename)
hx = 0.0
hy = 0.4
spirals = 10 #100
rays = 600   #6000
(xb,yb,zb,intensityb) = ln.zSpiralSpot(hx,hy,1,spirals,rays)
(xg,yg,zg,intensityg) = ln.zSpiralSpot(hx,hy,2,spirals,rays)
(xr,yr,zr,intensityr) = ln.zSpiralSpot(hx,hy,3,spirals,rays)
fig = plt.figure(facecolor='w')
ax = fig.add_subplot(111)
ax.set_aspect('equal')
ax.scatter(xr,yr,s=5,c='red',linewidth=0.35,zorder=20)
ax.scatter(xg,yg,s=5,c='lime',linewidth=0.35,zorder=21)
ax.scatter(xb,yb,s=5,c='blue',linewidth=0.35,zorder=22)
ax.set_xlabel('x');ax.set_ylabel('y')
fig.suptitle('Spiral Spot')
ax.grid(color='lightgray', linestyle='-', linewidth=1)
ax.ticklabel_format(scilimits=(-2,2))

# close the communication channel before calling show plot
pyz.closeLink()
plt.show()