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

if __name__ == '__main__':
    # Basic test
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
    ret = zArrayTrace(rd)
    print("ret = ", ret)

