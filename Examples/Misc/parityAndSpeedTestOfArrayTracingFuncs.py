# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:       parityAndSpeedTestOfArrayTracingFuncs.py
# Purpose:    1. parity tests
#                To verify the correctness of the array ray tracing functions by
#                checking the returned data against the existing single dde call
#                based ray tracing functions. The parity tests could potentially
#                be moved/used in the unit-tests for array tracing functions
#             2. speed tests
#                compute the average of best n execution times for array ray
#                tracing and compare that with single dde call based ray tracing
# Licence:    MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import time as time
import sys as sys
import pyzdde.arraytrace as at  # Module for array ray tracing
import pyzdde.zdde as pyz       # PyZDDE module
import os as os
from math import sqrt as sqrt

if sys.version_info > (3, 0):
    xrange = range

def set_up():
    """create dde link, and load lens into zemax, and push lens into
    zemax application (this is required for array ray tracing)
    """
    zmxfile = 'Cooke 40 degree field.zmx'
    ln = pyz.createLink()
    if not ln.zPushLensPermission():
        print("\nERROR: Extensions not allowed to push lenses. Please enable in Zemax.")
        ln.close()
        sys.exit(0)
    filename = os.path.join(ln.zGetPath()[1], "Sequential\Objectives", zmxfile)
    ln.zLoadFile(filename)
    ln.zGetUpdate()
    ln.zPushLens(1)
    return ln

def set_down(ln):
    """clean up, and close link
    """
    ln.zNewLens()
    ln.zPushLens(1)
    ln.close()

def no_error_in_ray_trace(rd, numRays):
    """This is not an "absolute". It assumes that if there have been no errors
    in tracing all rays, then the sum of all the "error" fields of the ray
    data array will be 0. However, it is possible, albeit with very small
    probability that few positive valued error may cancel out few negative
    valued errors.
    """
    sumErr = 0
    for i in range(1, numRays+1):
        sumErr += rd[i].error
    return sumErr == 0

def parity_zGetTrace_zArrayTrace_zGetTraceArray(ln, numRays):
    """function to check the parity between the ray traced data returned
    by zGetTrace(), zArrayTrace() and zGetTraceArray()
    """
    x, y = [0.0]*numRays, [0.0]*numRays
    l, m, n = [0.0]*numRays, [0.0]*numRays, [0.0]*numRays
    l2, m2, n2 = [0.0]*numRays, [0.0]*numRays, [0.0]*numRays
    hx, hy, mode, surf, waveNum = 0.0, 0.0, 0, -1, 1
    radius = int(sqrt(numRays)/2)
    k = 0
    for i in xrange(-radius, radius + 1, 1):
        for j in xrange(-radius, radius + 1, 1):
            px, py = i/(2*radius), j/(2*radius)
            tData = ln.zGetTrace(waveNum, mode, surf, hx, hy, px, py)
            _, _, x[k], y[k], _, l[k], m[k], n[k], l2[k], m2[k], n2[k], _, = tData
            k += 1
    # trace data from zArrayTrace
    _, rd = get_time_zArrayTrace(numRays, retRd=True)
    # compare the two ray traced data
    tol = 1e-10
    for k in range(numRays):
        assert abs(x[k] - rd[k+1].x) < tol, \
        "x[{}] = {}, rd[{}].x = {}".format(k, x[k], k+1, rd[k+1].x)
        assert abs(y[k] - rd[k+1].y) < tol, \
        "y[{}] = {}, rd[{}].y = {}".format(k, y[k], k+1, rd[k+1].y)
        assert abs(l[k] - rd[k+1].l) < tol, \
        "l[{}] = {}, rd[{}].l = {}".format(k, l[k], k+1, rd[k+1].l)
        assert abs(m[k] - rd[k+1].m) < tol, \
        "m[{}] = {}, rd[{}].m = {}".format(k, m[k], k+1, rd[k+1].m)
        assert abs(n[k] - rd[k+1].n) < tol, \
        "n[{}] = {}, rd[{}].n = {}".format(k, n[k], k+1, rd[k+1].n)
        assert abs(l2[k] - rd[k+1].Exr) < tol, \
        "l2[{}] = {}, rd[{}].l = {}".format(k, l2[k], k+1, rd[k+1].Exr)
        assert abs(m2[k] - rd[k+1].Eyr) < tol, \
        "m2[{}] = {}, rd[{}].m = {}".format(k, m2[k], k+1, rd[k+1].Eyr)
        assert abs(n2[k] - rd[k+1].Ezr) < tol, \
        "n2[{}] = {}, rd[{}].n = {}".format(k, n2[k], k+1, rd[k+1].Ezr)
    print("Parity test bw zGetTrace() & zArrayTrace() successful")
    # trace data from zGetTraceArray
    _, tData = get_time_zGetTraceArray(numRays, rettData=True)
    # compare the ray traced data
    for k in range(numRays):
        assert abs(x[k] - tData[2][k]) < tol, \
        "x[{}] = {}, tData[2][{}] = {}".format(k, x[k], k, tData[2][k])
        assert abs(y[k] - tData[3][k]) < tol, \
        "y[{}] = {}, tData[2][{}] = {}".format(k, y[k], k, tData[3][k])
        assert abs(l[k] - tData[5][k]) < tol, \
        "l[{}] = {}, tData[5][{}] = {}".format(k, l[k], k, tData[5][k])
        assert abs(m[k] - tData[6][k]) < tol, \
        "m[{}] = {}, tData[6][{}] = {}".format(k, m[k], k, tData[6][k])
        assert abs(n[k] - tData[7][k]) < tol, \
        "n[{}] = {}, tData[7][{}] = {}".format(k, n[k], k, tData[7][k])
        assert abs(l2[k] - tData[8][k]) < tol, \
        "l2[{}] = {}, tData[8][{}] = {}".format(k, l2[k], k, tData[8][k])
        assert abs(m2[k] - tData[9][k]) < tol, \
        "m2[{}] = {}, tData[9][{}] = {}".format(k, m2[k], k, tData[9][k])
        assert abs(n2[k] - tData[10][k]) < tol, \
        "n2[{}] = {}, tData[10][{}] = {}".format(k, n2[k], k, tData[10][k])
    print("Parity test bw zGetTrace() & zGetTraceArray() successful")

def parity_zGetTraceDirect_zGetTraceDirectArray(ln, numRays):
    """function to check the parity between the ray traced data returned
    by zGetTraceDirect() and zGetTraceDirectArray()
    """
    # use zGetTraceArray to surface # 2 to get ray coordinates and direction
    # cosines at surface 2
    radius = int(sqrt(numRays)/2)
    flatGrid = [(x/(2*radius),y/(2*radius)) for x in xrange(-radius, radius + 1, 1)
                      for y in xrange(-radius, radius + 1, 1)]
    px = [e[0] for e in flatGrid]
    py = [e[1] for e in flatGrid]
    tData0 = at.zGetTraceArray(numRays=numRays, px=px, py=py, waveNum=1, surf=2)
    assert sum(tData0[0]) == 0

    # use zGetTraceDirectArray to trace rays to the image surface using the
    # the ray coordinates and direction cosines at surface 2
    tData1 = at.zGetTraceDirectArray(numRays=numRays, x=tData0[2], y=tData0[3],
                                     z=tData0[4], l=tData0[5], m=tData0[6],
                                     n=tData0[7], waveNum=1, mode=0,
                                     startSurf=2, lastSurf=-1)
    assert sum(tData1[0]) == 0
    tol = 1e-10
    # use zGetTraceDirect to trace single rays per dde call
    for i in range(numRays):
        tData2 = ln.zGetTraceDirect(waveNum=1, mode=0, startSurf=2, stopSurf=-1,
                                    x=tData0[2][i], y=tData0[3][i], z=tData0[4][i],
                                    l=tData0[5][i], m=tData0[6][i], n=tData0[7][i])
        assert tData2[0] == 0
        for k in [2, 3, 4, 5, 6, 7]:
            assert abs(tData2[k] - tData1[k][i]) < tol, \
            ("tData2[{}] = {}, tData1[{}][{}] = {}"
            .format(k, tData2[k], k, i, tData1[k][i]))
    print("Parity test bw zGetTraceDirect() & zGetTraceDirectArray() successful")

def parity_zGetPolTrace_zGetPolTraceArray(ln, numRays):
    """function to check the parity between the ray traced data returned
    by zGetPolTrace() and zGetPolTraceArray()
    """
    radius = int(sqrt(numRays)/2)
    flatGrid = [(x/(2*radius),y/(2*radius)) for x in xrange(-radius, radius + 1, 1)
                      for y in xrange(-radius, radius + 1, 1)]
    px = [e[0] for e in flatGrid]
    py = [e[1] for e in flatGrid]
    polRtArrData = at.zGetPolTraceArray(numRays=numRays, px=px, py=py, Ey=1.0,
                                        waveNum=1, mode=0, surf=-1)
    assert sum(polRtArrData[0]) == 0
    tol = 1e-10
    # Trace rays using single DDE call
    for i in range(numRays):
        polRtData = ln.zGetPolTrace(waveNum=1, mode=0, surf=-1, hx=0, hy=0,
                                    px=px[i], py=py[i], Ex=0, Ey=1, Phax=0, Phay=0)
        assert polRtData.error == 0
        assert abs(polRtData.intensity - polRtArrData[1][i]) < tol, \
        ("polRtData.intensity = {},  polRtArrData[1][{}] = {}"
        .format(polRtData.intensity, i, polRtArrData[1][i]))
        assert abs(polRtData.Exr - polRtArrData[2][i]) < tol, \
        ("polRtData.Exr = {},  polRtArrData[2][{}] = {}"
        .format(polRtData.Exr, i, polRtArrData[2][i]))
        assert abs(polRtData.Exi - polRtArrData[3][i]) < tol, \
        ("polRtData.Exi = {},  polRtArrData[3][{}] = {}"
        .format(polRtData.Exi, i, polRtArrData[3][i]))
        assert abs(polRtData.Eyr - polRtArrData[4][i]) < tol, \
        ("polRtData.Eyr = {},  polRtArrData[4][{}] = {}"
        .format(polRtData.Eyr, i, polRtArrData[4][i]))
        assert abs(polRtData.Eyi - polRtArrData[5][i]) < tol, \
        ("polRtData.Eyi = {},  polRtArrData[5][{}] = {}"
        .format(polRtData.Eyi, i, polRtArrData[5][i]))
        assert abs(polRtData.Ezr - polRtArrData[6][i]) < tol, \
        ("polRtData.Ezr = {},  polRtArrData[6][{}] = {}"
        .format(polRtData.Ezr, i, polRtArrData[6][i]))
        assert abs(polRtData.Ezi - polRtArrData[7][i]) < tol, \
        ("polRtData.Ezi = {},  polRtArrData[7][{}] = {}"
        .format(polRtData.Ezi, i, polRtArrData[7][i]))
    print("Parity test bw zGetPolTrace() & zGetPolTraceArray() successful")

def parity_zGetPolTraceDirect_zGetPolTraceDirectArray(ln, numRays):
    """function to check the parity between the ray traced data returned
    by zGetPolTraceDirect() and zGetPolTraceDirectArray()
    """
    # use zGetTraceArray to surface # 2 to get ray coordinates and direction
    # cosines at surface 2
    radius = int(sqrt(numRays)/2)
    flatGrid = [(x/(2*radius),y/(2*radius)) for x in xrange(-radius, radius + 1, 1)
                      for y in xrange(-radius, radius + 1, 1)]
    px = [e[0] for e in flatGrid]
    py = [e[1] for e in flatGrid]
    tData0 = at.zGetTraceArray(numRays=numRays, px=px, py=py, waveNum=1, surf=2)
    assert sum(tData0[0]) == 0

    # use zGetPolTraceDirectArray to trace rays to the image surface using the
    # the ray coordinates and direction cosines at surface 2
    ptData1 = at.zGetPolTraceDirectArray(numRays=numRays, x=tData0[2], y=tData0[3],
                                         z=tData0[4], l=tData0[5], m=tData0[6],
                                         n=tData0[7], Ey=1.0, waveNum=1, mode=0,
                                         startSurf=2, lastSurf=-1)
    assert sum(ptData1[0]) == 0
    tol = 1e-10
    # Trace rays using single DDE call
    for i in range(numRays):
        ptData2 = ln.zGetPolTraceDirect(waveNum=1, mode=0, startSurf=2, stopSurf=-1,
                                        x=tData0[2][i], y=tData0[3][i], z=tData0[4][i],
                                        l=tData0[5][i], m=tData0[6][i], n=tData0[7][i],
                                        Ex=0, Ey=1.0, Phax=0, Phay=0)
        assert ptData2.error == 0
        assert (ptData2.intensity - ptData1[1][i]) < tol, \
        ("ptData2.intensity = {}, ptData1[1][{}] = {}"
        .format(ptData2.intensity, i, ptData1[1][i]))
        assert (ptData2.Exr - ptData1[2][i]) < tol, \
        ("ptData2.Exr = {}, ptData1[2][{}] = {}"
        .format(ptData2.Exr, i, ptData1[2][i]))
        assert (ptData2.Exi - ptData1[3][i]) < tol, \
        ("ptData2.Exi = {}, ptData1[3][{}] = {}"
        .format(ptData2.Exi, i, ptData1[3][i]))
        assert (ptData2.Eyr - ptData1[4][i]) < tol, \
        ("ptData2.Eyr = {}, ptData1[4][{}] = {}"
        .format(ptData2.Eyr, i, ptData1[4][i]))
        assert (ptData2.Eyi - ptData1[5][i]) < tol, \
        ("ptData2.Eyi = {}, ptData1[5][{}] = {}"
        .format(ptData2.Eyi, i, ptData1[5][i]))
        assert (ptData2.Ezr - ptData1[6][i]) < tol, \
        ("ptData2.Ezr = {}, ptData1[6][{}] = {}"
        .format(ptData2.Ezr, i, ptData1[6][i]))
        assert (ptData2.Ezi - ptData1[7][i]) < tol, \
        ("ptData2.Ezi = {}, ptData1[7][{}] = {}"
        .format(ptData2.Ezi, i, ptData1[7][i]))
    print("Parity test bw zGetPolTraceDirect() & zGetPolTraceDirectArray() successful")

def get_time_zArrayTrace(numRays, retRd=False):
    """return the time taken to perform ray tracing for the given number of rays
    using zArrayTrace() function.
    """
    radius = int(sqrt(numRays)/2)
    startTime = time.clock()
    rd = at.getRayDataArray(numRays, tType=0, mode=0)
    # Fill the rest of the ray data array,
    # hx, hy are zeros; mode = 0 (real), surf =  img surf, waveNum = 1
    k = 0
    for i in xrange(-radius, radius + 1, 1):
        for j in xrange(-radius, radius + 1, 1):
            k += 1
            rd[k].z = i/(2*radius)      # px
            rd[k].l = j/(2*radius)      # py
            rd[k].intensity = 1.0
            rd[k].wave = 1
    # Trace the rays
    ret = at.zArrayTrace(rd, timeout=5000)
    endTime = time.clock()
    if ret == 0 and no_error_in_ray_trace(rd, numRays):
        if retRd:
            return (endTime - startTime)*10e3, rd
        else:
            return (endTime - startTime)*10e3   # time in milliseconds

def get_time_zGetTraceArray(numRays, rettData=False):
    """return the time taken to perform tracing for the given number of rays
    using zGetTraceArray() function.
    """
    radius = int(sqrt(numRays)/2)
    startTime = time.clock()
    flatGrid = [(x/(2*radius),y/(2*radius)) for x in xrange(-radius, radius + 1, 1)
                      for y in xrange(-radius, radius + 1, 1)]
    px = [e[0] for e in flatGrid]
    py = [e[1] for e in flatGrid]
    tData = at.zGetTraceArray(numRays=numRays, px=px, py=py, waveNum=1)
    endTime = time.clock()
    if tData not in [-1, -999, -998] and sum(tData[0])==0: # tData[0] == error
        if rettData:
            return (endTime - startTime)*10e3, tData
        else:
            return (endTime - startTime)*10e3

def get_time_zGetTrace(ln, numRays):
    """return the time required to trace ``numRays`` number of rays
    using the PyZDDE function zGetTrace() that makes DDE call per trace
    """
    hx, hy, mode, surf, waveNum = 0.0, 0.0, 0, -1, 1
    radius = int(sqrt(numRays)/2)
    startTime = time.clock()
    errSum = 0
    for i in xrange(-radius, radius + 1, 1):
        for j in xrange(-radius, radius + 1, 1):
            px, py = i/(2*radius), j/(2*radius)
            errSum += ln.zGetTrace(waveNum, mode, surf, hx, hy, px, py)[0]
    endTime = time.clock()
    assert errSum == 0   # TO DO:: replace assert with something more meaningful
    return (endTime - startTime)*10e3   # time in milliseconds

def get_best_of_n_avg(seq, n=3):
    """compute the average of first n numbers in the list ``seq``
    sorted in ascending order
    """
    return sum(sorted(seq)[:n])/n

def compute_best_of_n_execution_times(func, numRays, numRuns, n, ln=None):
    """compute average execution time of func()
    numRays : list of the number of rays to trace
    numRuns : integer, number of times to execute the function ``func``
    n : integer, specifies how many execution times to average from the sorted
        list of execution times
    ln : PyZDDE link (only required for calling the single ray tracing
         functions)
    """
    bestnExecTimes = []
    funcName = (func.__name__).split('get_time_')[1]
    for nrays in numRays:
        print("Tracing {} rays {} times using {}".format(nrays, numRuns, funcName))
        execTimes = [0.0]*numRuns
        for i in range(numRuns):
            execTimes[i] = func(ln, nrays) if ln else func(nrays)
        bestnExecTimes.append(get_best_of_n_avg(execTimes, n))
    print("Average of best {} execution times = \n".format(n), bestnExecTimes)
    print("\n")

def speedtest_zGetTrace_zArrayTrace_zGetTraceArray(ln):
    print("\n")
    # i's must be odd, such that i**2, which is the number of rays to plot
    # is also odd
    numRays = [i**2 for i in xrange(3, 104, 10)]
    numRuns = 50
    n = 20
    # compute average of best of 10 execution times of zArrayTrace
    compute_best_of_n_execution_times(get_time_zArrayTrace, numRays, numRuns, n)
    # compute average execution time of zGetTraceArray()
    compute_best_of_n_execution_times(get_time_zGetTraceArray, numRays, numRuns, n)
    # compute average execution time of zGetTrace
    numRuns = 10
    n = 5
    compute_best_of_n_execution_times(get_time_zGetTrace, numRays, numRuns, n, ln)


if __name__ == '__main__':
    ln = set_up()
    # parity tests
    parity_zGetTrace_zArrayTrace_zGetTraceArray(ln, 81)
    parity_zGetTraceDirect_zGetTraceDirectArray(ln, 81)
    parity_zGetPolTrace_zGetPolTraceArray(ln, 81)
    parity_zGetPolTraceDirect_zGetPolTraceDirectArray(ln, 81)
    # speed test
    speedtest_zGetTrace_zArrayTrace_zGetTraceArray(ln)
    set_down(ln)