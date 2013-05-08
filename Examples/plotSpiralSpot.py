#-------------------------------------------------------------------------------
# Name:        plotSpiralSpot.py
# Purpose:     Example of using the "spiral spot" member function of pyZDDE.
#              Please note that this code uses matplotlib plotting library from
#              http://matplotlib.org/ for 2D-plotting
#
# Author:      Indranil Sinharoy
#
# Created:     06/10/2012
# Copyright:   (c) 2012, 2013
# Licence:     MIT License
#-------------------------------------------------------------------------------

import sys, os
import matplotlib.pyplot as plt

cd = os.getcwd()
ind = cd.find('Examples')
cd = cd[0:ind-1]

if cd not in sys.path:
    sys.path.append(cd)

from pyZDDE import *

# The ZEMAX file path
zmxfp = cd+'\\ZMXFILES\\'
zmxfile = 'Cooke 40 degree field.zmx'
filename = zmxfp+zmxfile

# Create a pyZDDE object
link0 = pyzdde()

# Initiate the DDE link
status = link0.zDDEInit()
if ~status:
    # Load a lens file into the ZEMAX DDE server
    ret = link0.zLoadFile(filename)
    if ~ret:
        status = link0.zPushLensPermission()
        if status:
            ret = link0.zPushLens(updateFlag=1)
            if ~ret:
                hx = 0.4
                hy = 0.0
                spirals = 100
                rays = 6000
                [xb,yb,zb,intensityb] = link0.spiralSpot(hx,hy,1,spirals,rays)
                [xg,yg,zg,intensityg] = link0.spiralSpot(hx,hy,2,spirals,rays)
                [xr,yr,zr,intensityr] = link0.spiralSpot(hx,hy,3,spirals,rays)
                fig,ax = plt.subplots(1,1)
                ax.set_aspect('equal')
                ax.scatter(xr,yr,s=8,c='red',linewidth=0.5,zorder=20)
                ax.scatter(xg,yg,s=8,c='lime',linewidth=0.5,zorder=21)
                ax.scatter(xb,yb,s=8,c='blue',linewidth=0.5,zorder=22)
                ax.set_xlabel('x');ax.set_ylabel('y')
                fig.suptitle('Spiral Spot')
                plt.show()
            else:
                print "Failed to push lens"
        else:
            print "Extension not allowed to push lens. Enable push permission."
    else:
        print "Could not load lens file"
    # close the DDE channel
    link0.zDDEClose()

else:
    print "DDE link could not be established"


