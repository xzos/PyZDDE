# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        arrayTraceSimple.py
# Purpose:     A simple script to demonstrate array ray tracing. We need to use
#              the two functions from the module pyzdde.arraytrace --
#              The first function getRayDataArray() helps us to create the
#              ray data structure array, and the function zArrayTrace() sends
#              the ray data to Zemax (through c) for tracing.
#
# Author:      Indranil Sinharoy
#
# Created:     Tue Feb 17 15:58:13 2015
# Copyright:   (c) Indranil Sinharoy, 2012 - 2015
# Licence:     MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import time as time
import pyzdde.arraytrace as at  # Module for array ray tracing
import pyzdde.zdde as pyz
import os as os

# The ZEMAX file path
cd = os.path.dirname(os.path.realpath(__file__))
ind = cd.find('Examples')
pDir = cd[0:ind-1]
zmxfile = 'Cooke 40 degree field.zmx'
filename = os.path.join(pDir, 'ZMXFILES', zmxfile)

def trace_rays():
    ln = pyz.createLink()
    ln.zLoadFile(filename)
    print("Loaded zemax file:", ln.zGetFile())
    ln.zGetUpdate()   # In general this should be done ...
    ln.zPushLens(1)   # FOR SOME REASON, THE ARRAY RAY TRACING SEEMS TO
                      # BE WORKING ON THE LENS THAT IS IN THE MAIN ZEMAX APPLICATION WINDOW!!!!
    ln.zNewLens()     # THIS IS JUST TO PROVE THE ABOVE POINT!!! RAY TRACING STILL ON THE LENS
                      # IN THE MAIN ZEMAX APPLICATION, EVENTHOUGH THE LENS IN THE DDE SERVER IS A "NEW LENS"
    numRays = 10201
    rd = at.getRayDataArray(numRays, tType=0, mode=0)

    # Fill the rest of the ray data array
    k = 0
    for i in xrange(-50, 51, 1):
        for j in xrange(-50, 51, 1):
            k += 1
            rd[k].z = i/100                   # px
            rd[k].l = j/100                   # py
            rd[k].intensity = 1.0
            rd[k].wave = 1

    # Trace the rays
    start_time = time.clock()
    ret = at.zArrayTrace(rd, timeout=5000)
    end_time = time.clock()
    print("Return value from array ray tracing:", ret)
    print("Ray tracing took", (end_time - start_time)*10e3, " milli seconds")

    # Dump the ray trace data into a file
    outputfile = os.path.join(cd, "arrayTraceOutput.txt")
    if ret==0:
        k = 0
        with open(outputfile, 'w') as f:
            f.write("Listing of Array trace data\n")
            f.write("     px      py error            xout            yout   trans\n")
            for i in xrange(-50, 51, 1):
                for j in xrange(-50, 51, 1):
                    k += 1
                    line = ("{:7.3f} {:7.3f} {:5d} {:15.6E} {:15.6E} {:7.4f}\n"
                            .format(i/100, j/100, rd[k].error, rd[k].x, rd[k].y, rd[k].intensity))
                    f.write(line)
        print("Success")
    else:
        print("There was some problem in ray tracing")

    ln.zNewLens()
    ln.zPushLens()
    ln.close()

if __name__ == '__main__':
    trace_rays()