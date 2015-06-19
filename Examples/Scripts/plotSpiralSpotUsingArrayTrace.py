# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:    plotSpiralSpotUsingArrayTrace.py
# Purpose: Demonstrate array tracing function by plotting spiral spot diagram
#          It also shows two ways of using the arraytrace module functions --
#          1. spiralSpot_using_zArrayTrace() uses the basic functions in which 
#             the user is responsible for (partially) creating the ray data 
#             structure before calling the array tracing function. Generally,
#             this option gives the fastest speed performance
#          2. spiralSpot_using_zGetTraceArray() uses the helper function 
#             zGetTraceArray() in arraytrace module that provides an easy 
#             interface to the user hiding away the details of construction of 
#             ray data structure and retrieval of traced data from the data
#             structure returned by Zemax. Generally this method is slightly
#             slower than 1, but very fast compared to tracing a ray per DDE
#             call. 
#             
#
# Author: Indranil Sinharoy
# Created: Apr 04, 2015
# License: MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import os.path as ospath
import matplotlib.pyplot as plt
import time as time
from math import pi, cos, sin
import numpy as np
import pyzdde.zdde as pyz
import pyzdde.arraytrace as at


def plotTracedData(x, y):
    """basic ray scatter plotting function. `x` and `y` are the ray intersection
    locations in the image surface
    """
    fig, ax = plt.subplots(1, 1, facecolor='w')
    ax.set_aspect('equal')
    ax.scatter(x, y, s=5, c='#0099FF', linewidth=0.25)
    ax.set_xlabel('X', fontsize=13)
    ax.set_ylabel('Y', fontsize=13)
    ax.set_title('Spiral Spot', fontsize=14)
    ax.ticklabel_format(scilimits=(-2,2))
    plt.show()


def trace_zSpiralSpot(ln, hx=0.0, hy=0.4, waveNum=1, spirals=10, numRays=600):
    """function to trace ray using the ``pyzdde.zdde`` function ``zSpiralSpot()``.
    
    This function traces one ray per DDE call
    """
    startTime = time.clock()
    x, y, z, _ = ln.zSpiralSpot(hx, hy, waveNum=waveNum, spirals=spirals, 
                                rays=numRays)
    endTime = time.clock()
    print("Execution time (zSpiralSpot) = {:4.2f}".format((endTime - startTime)*10e3), "ms")
    plotTracedData(x, y)


def spiralSpot_using_zArrayTrace(hx=0.0, hy=0.4, waveNum=1, spirals=10, numRays=600):
    """function replicates ``zSpiralSpot()`` using the ``pyzdde.arraytrace`` module 
    functions ``getRayDataArray()`` and ``zArrayTrace()``
    """
    startTime = time.clock()
    deltaTheta = (spirals*2.0*pi)/(numRays-1)
    # construct the ray data structure
    rd = at.getRayDataArray(numRays, tType=0, mode=0, endSurf=-1)
    for i in range(0, numRays):
        theta = i*deltaTheta
        r = i/(numRays-1)    
        px, py = r*cos(theta), r*sin(theta)
        rd[i+1].x = hx
        rd[i+1].y = hy
        rd[i+1].z = px
        rd[i+1].l = py
        rd[i+1].wave = waveNum
    # send ray data structure to Zemax for performing array tracing
    ret = at.zArrayTrace(rd)
    # retrieve the traced data from the ray data structure
    x, y = [0.0]*numRays, [0.0]*numRays
    if ret==0:                
        for i in range(1, numRays+1):
            x[i-1] = rd[i].x
            y[i-1] = rd[i].y
        endTime = time.clock()
        print("Execution time (zArrayTrace) = {:4.2f}".format((endTime - startTime)*10e3), "ms")
        plotTracedData(x, y)
    else:
        print("Error in tracing rays")

def spiralSpot_using_zGetTraceArray(hx=0.0, hy=0.4, waveNum=1, spirals=10, numRays=600):
    """function replicates ``zSpiralSpot()`` using the ``pyzdde.arraytrace`` module 
    helper function ``zGetTraceArray()`` 
    """
    startTime = time.clock()
    # create the hx, hy, px, py grid
    r = np.linspace(0, 1, numRays)
    theta = np.linspace(0, spirals*2.0*pi, numRays)
    px = (r*np.cos(theta)).tolist()
    py = (r*np.sin(theta)).tolist()
    # trace the rays
    tData = at.zGetTraceArray(numRays, [hx]*numRays, [hy]*numRays, px, py, 
                              waveNum=waveNum)
    # parse traced data and plot 
    err, _, x, y, _, _, _, _, _, _, _, _, _ = tData
    endTime = time.clock()
    print("Execution time (zGetTraceArray) = {:4.2f}".format((endTime - startTime)*10e3), "ms")
    if sum(err)==0:
        plotTracedData(x, y)
    else:
        print("Error in tracing rays")


if __name__=='__main__':
    # Create link and load Zemax file
    ln = pyz.createLink()
    filename = ospath.join(ln.zGetPath()[1], 'Sequential', 'Objectives', 
                           'Cooke 40 degree field.zmx')
    # Load a lens file into the Zemax DDE server
    ln.zLoadFile(filename)
    SPIRALS, RAYS = 100, 6000
    # spiral spot using pyzdde.zdde function that traces a ray per DDE call
    trace_zSpiralSpot(ln, spirals=SPIRALS, numRays=RAYS)
    # spiral spot using array tracing functions
    ln.zPushLens(1)
    spiralSpot_using_zArrayTrace(spirals=SPIRALS, numRays=RAYS)
    spiralSpot_using_zGetTraceArray(spirals=SPIRALS, numRays=RAYS)
    # clean up the LDE and close DDE link
    ln.zNewLens()
    pyz.closeLink()