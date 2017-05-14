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
# Copyright:   (c) Indranil Sinharoy, 2012 - 2017
# Licence:     MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import pyzdde.arraytrace as at  # Module for array ray tracing
import pyzdde.zdde as pyz
import os as os
import sys as sys
from math import sqrt as sqrt

if sys.version_info[0] > 2:
    xrange = range
cd = os.path.dirname(os.path.realpath(__file__))

def trace_rays():
    ln = pyz.createLink()
    filename = os.path.join(ln.zGetPath()[1], 'Sequential', 'Objectives', 
                            'Cooke 40 degree field.zmx')
    ln.zLoadFile(filename)
    print("Loaded zemax file:", ln.zGetFile())
    ln.zGetUpdate()   # In general this should be done ...
    if not ln.zPushLensPermission():
        print("\nERROR: Extensions not allowed to push lenses. Please enable in Zemax.")
        ln.close()
        sys.exit(0)
    ln.zPushLens(1)   # FOR SOME REASON, THE ARRAY RAY TRACING SEEMS TO
                      # BE WORKING ON THE LENS THAT IS IN THE MAIN ZEMAX APPLICATION WINDOW!!!!
    ln.zNewLens()     # THIS IS JUST TO PROVE THE ABOVE POINT!!! RAY TRACING STILL ON THE LENS
                      # IN THE MAIN ZEMAX APPLICATION, EVENTHOUGH THE LENS IN THE DDE SERVER IS A "NEW LENS"
    numRays = 101**2    # 10201
    rd = at.getRayDataArray(numRays, tType=0, mode=0, endSurf=-1)
    radius = int(sqrt(numRays)/2)

    # Fill the rest of the ray data array
    k = 0
    for i in xrange(-radius, radius + 1, 1):
        for j in xrange(-radius, radius + 1, 1):
            k += 1
            rd[k].z = i/(2*radius)                   # px
            rd[k].l = j/(2*radius)                   # py
            rd[k].intensity = 1.0
            rd[k].wave = 1

    # Trace the rays
    ret = at.zArrayTrace(rd, timeout=5000)

    # Dump the ray trace data into a file
    outputfile = os.path.join(cd, "arrayTraceOutput.txt")
    if ret==0:
        k = 0
        with open(outputfile, 'w') as f:
            f.write("Listing of Array trace data\n")
            f.write("     px      py error            xout            yout"
                    "         l         m         n    opd    Exr     Exi"
                    "     Eyr     Eyi     Ezr     Ezi    trans\n")
            for i in xrange(-radius, radius + 1, 1):
                for j in xrange(-radius, radius + 1, 1):
                    k += 1
                    line = ("{:7.3f} {:7.3f} {:5d} {:15.6E} {:15.6E} {:9.5f} "
                            "{:9.5f} {:9.5f} {:7.3f} {:7.3f} {:7.3f} {:7.3f} "
                            "{:7.3f} {:7.3f} {:7.3f} {:7.4f}\n"
                            .format(i/(2*radius), j/(2*radius), rd[k].error,
                                    rd[k].x, rd[k].y, rd[k].l, rd[k].m, rd[k].n,
                                    rd[k].opd, rd[k].Exr, rd[k].Exi, rd[k].Eyr,
                                    rd[k].Eyi, rd[k].Ezr, rd[k].Ezi, rd[k].intensity))
                    f.write(line)
        print("Success")
        print("Ray trace data outputted to the file {}".format(outputfile))
    else:
        print("There was some problem in ray tracing")

    ln.zNewLens()
    ln.zPushLens()
    ln.close()

if __name__ == '__main__':
    trace_rays()