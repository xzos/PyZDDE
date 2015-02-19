#-------------------------------------------------------------------------------
# Name:        arraytrace.py
# Purpose:     Module for doing array ray tracing in Zemax.
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2015
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.8.03
#-------------------------------------------------------------------------------
"""Module for doing array ray tracing in Zemax."""
from __future__ import print_function
import os as _os
import sys as _sys
import ctypes as _ct

# Ray data structure as defined in Zemax manual
class DdeArrayData(_ct.Structure):
    _fields_ = [(  'x', _ct.c_double), ('y', _ct.c_double), ('z', _ct.c_double),
                (  'l', _ct.c_double), ('m', _ct.c_double), ('n', _ct.c_double),
                ('opd', _ct.c_double), ('intensity', _ct.c_double),
                ('Exr', _ct.c_double), ('Exi', _ct.c_double),
                ('Eyr', _ct.c_double), ('Eyi', _ct.c_double),
                ('Ezr', _ct.c_double), ('Ezi', _ct.c_double),
                ('wave', _ct.c_int),   ('error', _ct.c_int),
                ('vigcode', _ct.c_int), ('want_opd', _ct.c_int)]

def _is64bit():
    """return True if Python version is 64 bit
    """
    return _sys.maxsize > 2**31 - 1

_dllDir = "arraytrace\\x64\\Release\\" if _is64bit() else "arraytrace\\Release\\"
_dllName = "ArrayTrace.dll"
_dllpath = _os.path.join(_os.path.dirname(_os.path.realpath(__file__)), _dllDir)
# load the arrayTrace library
_array_trace_lib = _ct.WinDLL(_dllpath + _dllName)
_arrayTrace = _array_trace_lib.arrayTrace
# specify argtypes and restype
_arrayTrace.restype = _ct.c_int
_arrayTrace.argtypes = [_ct.POINTER(DdeArrayData), _ct.c_uint]


def zArrayTrace(rd, timeout=5000):
    """function to trace large number of rays

    Parameters
    ----------
    rd : ctypes array
        array of ray data structure as specified in Zemax manual for array
        ray tracing. Use the helper function getRayDataArray() to generate ``rd``

    timeout : integer
        time in milliseconds (Default = 5000)

    Returns
    -------
    ret : integer
        Error codes meaning 0 = SUCCESS, -1 = Couldn't retrieve data in
        PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
        -998 = timeout reached
    """
    return _arrayTrace(rd, int(timeout))

def getRayDataArray(numRays, tType=0, mode=0, startSurf=None, endSurf=-1):
    """helper function to create the basic ray data array (rd). The caller
    must fill the appropriate fields of the array from elements rd[1] to
    rd[numRays]

    Parameters
    ----------
    numRays : integer
        number of rays that will be traced
    tType : integer (0-3)
        0 =  GetTrace (Default), 1 = GetTraceDirect, 2 = GetPolTrace,
        3 = GetPolTraceDirect
    mode : integer (0-1)
        0 = real (Default), 1 = paraxial
    startSurf : integer or None (Default)
        specify start surface in tType 1 and 3
    endSurf : integer
        specify end surface to trace. Default is image surface (-1)

    Returns
    -------
    rd : ctypes array
        array of ray data structure as specified in Zemax manual for array
        ray tracing.

    Notes
    -----
    Since the memory for the array is allocated by Python, the user doesn't
    need to worry about freeing the memory.
    """
    rd = (DdeArrayData * (numRays + 1))()
    # Setup a basic ray data array for test
    rd[0].opd = _ct.c_double(tType)
    rd[0].wave = _ct.c_int(mode)
    rd[0].error = _ct.c_int(numRays)
    if startSurf:
        rd[0].vigcode = _ct.c_int(startSurf)
    rd[0].want_opd = _ct.c_int(endSurf)
    return rd

def zGetTraceArray(numRays, hx=None, hy=None, px=None, py=None, intensity=None,
                   waveNum=None, mode=0, surf=-1, want_opd=0, timeout=5000):
    """Trace large number of rays defined by their normalized field and pupil
    coordinates.

    Parameters
    ----------
    numRays : integer
        number of rays to trace. ``numRays`` should be equal to the length
        of the lists (if provided) ``hx``, ``hy``, ``px``, etc.
    hx : list, optional
        list of normalized field heights along x axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hx` is created.
    hy : list, optional
        list of normalized field heights along y axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hy` is created
    px : list, optional
        list of normalized heights in pupil coordinates, along x axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``px` is created.
    py : list, optional
        list of normalized heights in pupil coordinates, along y axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``py` is created
    intensity : float or list, optional
        initial intensities. If a list of length ``numRays`` is given it is
        used. If a single float value is passed, all rays use the same value for
        their initial intensities. If ``None``, all rays use a value of ``1.0``
        as their initial intensities.
    waveNum : integer or list (of integers), optional
        wavelength number. If a list of integers of length ``numRays`` is given
        it is used. If a single integer value is passed, all rays use the same
        value for wavelength number. If ``None``, all rays use wavelength
        number equal to 1.
    mode : integer, optional
        0 = real (Default), 1 = paraxial
    surf : integer, optional
        surface to trace the ray to. Usually, the ray data is only needed at
        the image surface (``surf = -1``, default)
    want_opd : integer, optional
        0 if OPD data is not needed (Default), 1 if it is. See Zemax manual
        for details.
    timeout : integer, optional
        command timeout specified in milli-seconds

    Returns
    -------
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the direction cosines after refraction into the media following
        the requested surface.
    opd : list of reals
        computed optical path difference if ``want_opd > 0``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.
    Exr, Eyr, Ezr : list of reals
        list of x or y or z cosine of the surface normal
    error : list of integers
        0 = ray traced successfully;
        +ve number = the ray missed the surface;
        -ve number = the ray total internal reflected (TIR) at surface
                     given by the absolute value of the ``error``
    vigcode : list of integers
        the first surface where the ray was vignetted. Unless an error occurs
        at that surface or subsequent to that surface, the ray will continue
        to trace to the requested surface.

    Notes
    -----
    The opd can only be computed if the last surface is the image surface,
    otherwise, the opd value will be zero.
    """
    rd = getRayDataArray(numRays, tType=0, mode=mode, endSurf=surf)
    hx = hx if hx else [0.0] * numRays
    hy = hy if hy else [0.0] * numRays
    px = px if px else [0.0] * numRays
    py = py if py else [0.0] * numRays
    if intensity:
        intensity = intensity if isinstance(intensity, list) else [intensity]*numRays
    else:
        intensity = [1.0] * numRays
    if waveNum:
        waveNum = waveNum if isinstance(waveNum, list) else [waveNum]*numRays
    else:
        waveNum = [1] * numRays
    want_opd = [want_opd] * numRays
    print("Want_OPD = ", want_opd)
    # fill up the structure
    for i in xrange(1, numRays+1):
        rd[i].x = hx[i-1]
        rd[i].y = hy[i-1]
        rd[i].px = px[i-1]
        rd[i].py = py[i-1]
        rd[i].intensity = intensity[i-1]
        rd[i].wave = waveNum[i-1]
        rd[i].want_opd = want_opd[i-1]
    # call ray tracing
    ret = zArrayTrace(rd, timeout)
    if ret == 0:
        reals = ['x', 'y', 'z', 'l', 'm', 'n', 'opd',
                 'intensity', 'Exr', 'Eyr', 'Ezr']
        ints = ['error', 'vigcode']
        for r in reals:
            exec(r + " = [0.0] * numRays")
        for i in ints:
            exec(i + " = [0] * numRays")
        for i in xrange(1, numRays+1):
            x[i-1] = rd[i].x
            y[i-1] = rd[i].y
            z[i-1] = rd[i].z
            l[i-1] = rd[i].l
            m[i-1] = rd[i].m
            n[i-1] = rd[i].n
            opd[i-1] = rd[i].opd
            intensity[i-1] = rd[i].intensity
            Exr[i-1] = rd[i].Exr
            Eyr[i-1] = rd[i].Eyr
            Ezr[i-1] = rd[i].Ezr
            error[i-1] = rd[i].error
            vigcode[i-1] = rd[i].vigcode
        return x, y, z, l, m, n, opd, intensity, Exr, Eyr, Ezr, error, vigcode
    else:
        print("Error. zArrayTrace returned error code {}".format(ret))


if __name__ == '__main__':
    # Basic test
    print("Basic test of zArrayTrace:")
    nr = 441
    rd = getRayDataArray(nr)
    # Fill the rest of the ray data array
    k = 0
    for i in xrange(-10, 11, 1):
        for j in xrange(-10, 11, 1):
            k += 1
            rd[k].z = i/20.0                   # px
            rd[k].l = j/20.0                   # py
            rd[k].intensity = 1.0
            rd[k].wave = 1
            rd[k].want_opd = 0
    ret = zArrayTrace(rd)
    print("ret = ", ret)
    for i in range(1, 11):                     # SEEMS LIKE ZEMAX IS CALCULATING THE OPD EVEN WHEN NOT ASKED
        print(rd[k].opd)
    print("\nBasic test of zGetTraceArray:")
    hx = [(i - 5.0)/10.0 for i in range(11)]
    hy = [(i - 5.0)/10.0 for i in range(11)]
    ret = zGetTraceArray(numRays=len(hx), hx=hx, hy=hy)
    if ret not in []:
        x, y, z, l, m, n, opd, intensity, Exr, Eyr, Ezr, error, vigcode = ret
        print("x = ", x)
        print("y = ", y)
        print("z = ", z)
        print("l = ", l)
        print("opd = ", opd)
        print("intensity = ", intensity)
        print("Exr = ", Exr)
        print("err = ", error)
    else:
        print("ret = ", ret)

