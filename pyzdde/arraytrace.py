# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        arraytrace.py
# Purpose:     Module for doing array ray tracing in Zemax.
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
"""Module for doing array ray tracing as described in Zemax manual. This module
defines the DDE ray data structure using ctypes, and provides the following
two main functions:

    1. zArrayTrace() -- The main function for calling Zemax for array ray tracing
    2. getRayDataArray() -- Helper function that creates the ctypes ray data structure
                            array, fills up the first element and returns the array

In addition the following helper functions are provided that supports 5 different
modes discussed in the Zemax manual

    1. zGetTraceArray()
    2. zGetTraceDirectArray()
    3. zGetPolTraceArray()
    4. zGetPolTraceDirectArray()
    5. zGetNSCTraceArray()
"""
from __future__ import print_function
import os as _os
import sys as _sys
import ctypes as _ct
import collections as _co
#import gc as _gc

if _sys.version_info[0] > 2:
    xrange = range

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
    """function to trace large number of rays on lens file in the LDE of main
    Zemax application (not in the DDE server)

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

def getRayDataArray(numRays, tType=0, mode=0, startSurf=None, endSurf=-1,
                    **kwargs):
    """function to create the basic ray data structure array, ``rd``.

    The caller must fill the rest of the appropriate fields of the
    array from elements ``rd[1]`` to ``rd[numRays]``

    Parameters
    ----------
    numRays : integer
        number of rays that will be traced. ``numRays`` is used to determine
        the size of the ray data structure array, which is ``numRays + 1``. In
        addition, it also sets the ``error`` field of the ray data structure.
    tType : integer (0-3)
        0 =  GetTrace (Default), 1 = GetTraceDirect, 2 = GetPolTrace,
        3 = GetPolTraceDirect
        (``tType`` sets the "opd" field of the ray data structure)
    mode : integer (0-1)
        0 = real (Default), 1 = paraxial
        (``mode`` sets the ``wave`` field of the ray data structure)
    startSurf : integer or None (Default)
        specify start surface in ``tType`` 1 and 3.
        (``startSurf`` sets the ``vigcode`` field of the ray data structure)
    endSurf : integer
        specify end surface to trace. Default is image surface (-1).
        (``endSurf`` sets the ``want_opd`` field of the ray data structure)
    kwargs : keyword arguments
        The ``kwargs`` are used to set the fields of the 0th element using the
        field names as specified in the Zemax manual. For example, use
        ``x``, ``y``, ``z``, and ``l`` for ``Ex``, ``Ey``, ``Phax``, and
        ``Phay`` in ``zGetPolTraceArray()`` and ``zGetPolTraceDirectArray()``
        (``tType=2``).
        In case if a particular field value is specified by both the regular
        parameters and by ``kwargs``, the final value set is the value
        provided by ``kwargs``. For example, the ``numRays`` is used to
        determine the length of the array, and set the ``error`` field
        of the 0th element. However, calling
        ``getRayDataArray(numRays=100, error=1)`` will create an array of
        length 101, and set the ``error`` field of the 0th element to ``1``

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
    fields = {'x', 'y', 'z', 'l', 'm', 'n', 'opd', 'intensity', 'Exr', 'Exi',
              'Eyr', 'Eyi', 'Ezr', 'Ezi', 'wave', 'error', 'vigcode', 'want_opd'}
    if kwargs:
        assert set(kwargs).issubset(fields), "Received one or more unexpected kwargs "
    # create ctypes array
    rd = (DdeArrayData * (numRays + 1))()
    # Setup a basic ray data array for test
    rd[0].opd = _ct.c_double(tType)
    rd[0].wave = _ct.c_int(mode)
    rd[0].error = _ct.c_int(numRays)
    if startSurf:
        rd[0].vigcode = _ct.c_int(startSurf)
    rd[0].want_opd = _ct.c_int(endSurf)
    # fill up based on kwargs. This will also override any previously set
    # fields
    if kwargs:
        for k in kwargs:
            setattr(rd[0], k, kwargs[k])
    return rd

def zGetTraceArray(numRays, hx=None, hy=None, px=None, py=None, intensity=None,
                   waveNum=None, mode=0, surf=-1, want_opd=0, timeout=5000):
    """Trace large number of rays defined by their normalized field and pupil
    coordinates on lens file in the LDE of main Zemax application (not in the DDE server)

    Parameters
    ----------
    numRays : integer
        number of rays to trace. ``numRays`` should be equal to the length
        of the lists (if provided) ``hx``, ``hy``, ``px``, etc.
    hx : list, optional
        list of normalized field heights along x axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hx`` is created.
    hy : list, optional
        list of normalized field heights along y axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hy`` is created
    px : list, optional
        list of normalized heights in pupil coordinates, along x axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``px`` is created.
    py : list, optional
        list of normalized heights in pupil coordinates, along y axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``py`` is created
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
    error : list of integers
        0 = ray traced successfully;
        +ve number = the ray missed the surface;
        -ve number = the ray total internal reflected (TIR) at surface
                     given by the absolute value of the ``error``
    vigcode : list of integers
        the first surface where the ray was vignetted. Unless an error occurs
        at that surface or subsequent to that surface, the ray will continue
        to trace to the requested surface.
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the x, y, and z direction cosines after refraction into the media
        following the requested surface.
    l2, m2, n2 : list of reals
        list of x or y or z surface intercept direction normals at requested
        surface
    opd : list of reals
        computed optical path difference if ``want_opd > 0``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.

    If ray tracing fails, a single integer error code is returned,
    which has the following meaning: -1 = Couldn't retrieve data in
    PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
    -998 = timeout reached

    Examples
    -------- 
    >>> n = 9**2
    >>> nx = np.linspace(-1, 1, np.sqrt(n))
    >>> hx, hy = np.meshgrid(nx, nx)
    >>> hx, hy = hx.flatten().tolist(), hy.flatten().tolist()
    >>> rayData = at.zGetTraceArray(numRays=n, hx=hx, hy=hy, mode=0)
    >>> err, vig = rayData[0], rayData[1]
    >>> x, y, z = rayData[2], rayData[3], rayData[4]

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
    # fill up the structure
    for i in xrange(1, numRays+1):
        rd[i].x = hx[i-1]
        rd[i].y = hy[i-1]
        rd[i].z = px[i-1]
        rd[i].l = py[i-1]
        rd[i].intensity = intensity[i-1]
        rd[i].wave = waveNum[i-1]
        rd[i].want_opd = want_opd[i-1]

    # call ray tracing
    ret = zArrayTrace(rd, timeout)

    # free up some memory
    #del hx, hy, px, py, intensity, waveNum, want_opd # seems to increase running time
    #_gc.collect()
    d = {}
    if ret == 0:
        reals = ['x', 'y', 'z', 'l', 'm', 'n', 'l2', 'm2', 'n2', 'opd',
                 'intensity']
        ints = ['error', 'vigcode']
        for r in reals:
            exec(r + " = [0.0] * numRays", locals(), d)
        for i in ints:
            exec(i + " = [0] * numRays", locals(), d)
        for i in xrange(1, numRays+1):
            d["x"][i-1] = rd[i].x
            d["y"][i-1] = rd[i].y
            d["z"][i-1] = rd[i].z
            d["l"][i-1] = rd[i].l
            d["m"][i-1] = rd[i].m
            d["n"][i-1] = rd[i].n
            d["opd"][i-1] = rd[i].opd
            d["intensity"][i-1] = rd[i].intensity
            d["l2"][i-1] = rd[i].Exr
            d["m2"][i-1] = rd[i].Eyr
            d["n2"][i-1] = rd[i].Ezr
            d["error"][i-1] = rd[i].error
            d["vigcode"][i-1] = rd[i].vigcode
        return (d["error"], d["vigcode"], d["x"], d["y"], d["z"], 
                d["l"], d["m"], d["n"], d["l2"], d["m2"], d["n2"], 
                d["opd"], d["intensity"])
    else:
        return ret

def zGetTraceDirectArray(numRays, x=None, y=None, z=None, l=None, m=None,
                         n=None, intensity=None, waveNum=None, mode=0,
                         startSurf=0, lastSurf=-1, timeout=5000):
    """Trace large number of rays defined by ``x``, ``y``, ``z``, ``l``,
    ``m`` and ``n`` coordinates on any starting surface as well as
    wavelength number, mode and the surface to trace the ray to. 
    
    Ray tracing is performed on the lens file in the LDE of main Zemax 
    application (not in the DDE server)

    Parameters
    ----------
    numRays : integer
        number of rays to trace. ``numRays`` should be equal to the length
        of the lists (if provided) ``x``, ``y``, ``x``, etc.
    x : list, optional
        list specifying the x coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``x`` is created.
    y : list, optional
        list specifying the y coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``y`` is created
    z : list, optional
        list specifying the z coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``z`` is created.
    l : list, optional
        list of x-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``l`` is created
    m : list, optional
        list of y-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``m`` is created
    n : list, optional
        list of z-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``n`` is created
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
    startSurf : integer, optional
        start surface number (default = 0)
    lastSurf : integer, optional
        surface to trace the ray to, which is any valid surface number
        (``surf = -1``, default)
    timeout : integer, optional
        command timeout specified in milli-seconds

    Returns
    -------
    error : list of integers
        0 = ray traced successfully;
        +ve number = the ray missed the surface;
        -ve number = the ray total internal reflected (TIR) at surface
                     given by the absolute value of the ``error``
    vigcode : list of integers
        the first surface where the ray was vignetted. Unless an error occurs
        at that surface or subsequent to that surface, the ray will continue
        to trace to the requested surface.
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the x, y, and z direction cosines after refraction into the media
        following the requested surface.
    l2, m2, n2 : list of reals
        list of x or y or z surface intercept direction normals at
        requested surface
    opd : list of reals
        computed optical path difference if ``want_opd > 0``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.

    If ray tracing fails, a single integer error code is returned,
    which has the following meaning: -1 = Couldn't retrieve data in
    PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
    -998 = timeout reached

    Notes
    -----
    Computation of OPD is not permitted in this mode.
    """
    rd = getRayDataArray(numRays, tType=1, mode=mode, startSurf=startSurf,
                         endSurf=lastSurf)
    x = x if x else [0.0] * numRays
    y = y if y else [0.0] * numRays
    z = z if z else [0.0] * numRays
    l = l if l else [0.0] * numRays
    m = m if m else [0.0] * numRays
    n = n if n else [0.0] * numRays
    if intensity:
        intensity = intensity if isinstance(intensity, list) else [intensity]*numRays
    else:
        intensity = [1.0] * numRays
    if waveNum:
        waveNum = waveNum if isinstance(waveNum, list) else [waveNum]*numRays
    else:
        waveNum = [1] * numRays

    # fill up the structure
    for i in xrange(1, numRays+1):
        rd[i].x = x[i-1]
        rd[i].y = y[i-1]
        rd[i].z = z[i-1]
        rd[i].l = l[i-1]
        rd[i].m = m[i-1]
        rd[i].n = n[i-1]
        rd[i].intensity = intensity[i-1]
        rd[i].wave = waveNum[i-1]

    # call ray tracing
    ret = zArrayTrace(rd, timeout)
    d = {}
    if ret == 0:
        reals = ['x', 'y', 'z', 'l', 'm', 'n', 'l2', 'm2', 'n2', 'opd',
                 'intensity']
        ints = ['error', 'vigcode']
        for r in reals:
            exec(r + " = [0.0] * numRays", locals(), d)
        for i in ints:
            exec(i + " = [0] * numRays", locals(), d)
        for i in xrange(1, numRays+1):
            d["x"][i-1] = rd[i].x
            d["y"][i-1] = rd[i].y
            d["z"][i-1] = rd[i].z
            d["l"][i-1] = rd[i].l
            d["m"][i-1] = rd[i].m
            d["n"][i-1] = rd[i].n
            d["opd"][i-1] = rd[i].opd
            d["intensity"][i-1] = rd[i].intensity
            d["l2"][i-1] = rd[i].Exr
            d["m2"][i-1] = rd[i].Eyr
            d["n2"][i-1] = rd[i].Ezr
            d["error"][i-1] = rd[i].error
            d["vigcode"][i-1] = rd[i].vigcode
        return (d["error"], d["vigcode"], d["x"], d["y"], d["z"], 
                d["l"], d["m"], d["n"], d["l2"], d["m2"], d["n2"], 
                d["opd"], d["intensity"])
    else:
        return ret

def zGetPolTraceArray(numRays, hx=None, hy=None, px=None, py=None, Exr=None,
                      Exi=None, Eyr=None, Eyi=None, Ezr=None, Ezi=None, Ex=0,
                      Ey=0, Phax=0, Phay=0, intensity=None, waveNum=None, mode=0,
                      surf=-1, timeout=5000):
    """Trace large number of polarized rays defined by their normalized
    field and pupil coordinates. Similar to ``GetPolTrace()``
    
    Ray tracing is performed on the lens file in the LDE of main Zemax 
    application (not in the DDE server)

    Parameters
    ----------
    numRays : integer
        number of rays to trace. ``numRays`` should be equal to the length
        of the lists (if provided) ``hx``, ``hy``, ``px``, etc.
    hx : list, optional
        list of normalized field heights along x axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hx`` is created.
    hy : list, optional
        list of normalized field heights along y axis, of length ``numRays``;
        if ``None``, a list of 0.0s for ``hy`` is created
    px : list, optional
        list of normalized heights in pupil coordinates, along x axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``px`` is created.
    py : list, optional
        list of normalized heights in pupil coordinates, along y axis, of
        length ``numRays``; if ``None``, a list of 0.0s for ``py`` is created
    Exr : list, optional
        list of real part of the electric field in x direction for each ray.
        if ``None``, a list of 0.0s for ``Exr`` is created. See Notes
    Exi : list, optional
        list of imaginary part of the electric field in x direction for each ray.
        if ``None``, a list of 0.0s for ``Exi`` is created. See Notes
    Eyr : list, optional
        list of real part of the electric field in y direction for each ray.
        if ``None``, a list of 0.0s for ``Eyr`` is created. See Notes
    Eyi : list, optional
        list of imaginary part of the electric field in y direction for each ray.
        if ``None``, a list of 0.0s for ``Eyi`` is created. See Notes
    Ezr : list, optional
        list of real part of the electric field in z direction for each ray.
        if ``None``, a list of 0.0s for ``Ezr`` is created. See Notes
    Ezi : list, optional
        list of imaginary part of the electric field in z direction for each ray.
        if ``None``, a list of 0.0s for ``Ezi`` is created. See Notes
    Ex : float
        normalized electric field magnitude in x direction to be defined in
        array position 0. If not provided, an unpolarized ray will be traced.
        See Notes.
    Ey : float
        normalized electric field magnitude in y direction to be defined in
        array position 0. If not provided, an unpolarized ray will be traced.
        See Notes.
    Phax : float
        relative phase in x direction in degrees
    Phay : float
        relative phase in y direction in degrees
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
        surface to trace the ray to. (``surf = -1``, default)
    timeout : integer, optional
        command timeout specified in milli-seconds

    Returns
    -------
    error : list of integers
        0 = ray traced successfully;
        +ve number = the ray missed the surface;
        -ve number = the ray total internal reflected (TIR) at surface
                     given by the absolute value of the ``error``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.
    Exr : list of real values
        list of real parts of the electric field components in x
    Exi : list of real values
        list of imaginary parts of the electric field components in x
    Eyr : list of real values
        list of real parts of the electric field components in y
    Eyi : list of real values
        list of imaginary parts of the electric field components in y
    Ezr : list of real values
        list of real parts of the electric field components in z
    Ezi : list of real values
        list of imaginary parts of the electric field components in z

    If ray tracing fails, a single integer error code is returned,
    which has the following meaning: -1 = Couldn't retrieve data in
    PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
    -998 = timeout reached

    Notes
    -----
    1. If all six of the electric field values ``Exr``, ``Exi``, ``Eyr``,
       ``Eyi``, ``Ezr``, and ``Ezi`` for a ray are zero Zemax will use the
       ``Ex`` and ``Ey`` values provided in array position 0 to determine
       the electric field. Otherwise, the electric field is defined by these
       six values.
    2. The defined electric field vector must be orthogonal to the ray vector
       or incorrect ray tracing will result.
    3. Even if these six values are defined for each ray, values for ``Ex``
       and ``Ey`` in the array position 0 must still be defined, otherwise
       an unpolarized ray trace will result.
    """
    rd = getRayDataArray(numRays, tType=2, mode=mode, endSurf=surf,
                         x=Ex, y=Ey, z=Phax, l=Phay)
    hx = hx if hx else [0.0] * numRays
    hy = hy if hy else [0.0] * numRays
    px = px if px else [0.0] * numRays
    py = py if py else [0.0] * numRays
    Exr = Exr if Exr else [0.0] * numRays
    Exi = Exi if Exi else [0.0] * numRays
    Eyr = Eyr if Eyr else [0.0] * numRays
    Eyi = Eyi if Eyi else [0.0] * numRays
    Ezr = Ezr if Ezr else [0.0] * numRays
    Ezi = Ezi if Ezi else [0.0] * numRays

    if intensity:
        intensity = intensity if isinstance(intensity, list) else [intensity]*numRays
    else:
        intensity = [1.0] * numRays
    if waveNum:
        waveNum = waveNum if isinstance(waveNum, list) else [waveNum]*numRays
    else:
        waveNum = [1] * numRays

    # fill up the structure
    for i in xrange(1, numRays+1):
        rd[i].x = hx[i-1]
        rd[i].y = hy[i-1]
        rd[i].z = px[i-1]
        rd[i].l = py[i-1]
        rd[i].Exr = Exr[i-1]
        rd[i].Exi = Exi[i-1]
        rd[i].Eyr = Eyr[i-1]
        rd[i].Eyi = Eyi[i-1]
        rd[i].Ezr = Ezr[i-1]
        rd[i].Ezi = Ezi[i-1]
        rd[i].intensity = intensity[i-1]
        rd[i].wave = waveNum[i-1]

    # call ray tracing
    ret = zArrayTrace(rd, timeout)
    d = {}
    if ret == 0:
        reals = ['intensity', 'Exr', 'Exi', 'Eyr', 'Eyi', 'Ezr', 'Ezi']
        ints = ['error', ]
        for r in reals:
            exec(r + " = [0.0] * numRays", locals(), d)
        for i in ints:
            exec(i + " = [0] * numRays", locals(), d)
        for i in xrange(1, numRays+1):
            d["intensity"][i-1] = rd[i].intensity
            d["Exr"][i-1] = rd[i].Exr
            d["Exi"][i-1] = rd[i].Exi
            d["Eyr"][i-1] = rd[i].Eyr
            d["Eyi"][i-1] = rd[i].Eyi
            d["Ezr"][i-1] = rd[i].Ezr
            d["Ezi"][i-1] = rd[i].Ezi
            d["error"][i-1] = rd[i].error
        return (d["error"], d["intensity"], 
                d["Exr"], d["Exi"], d["Eyr"], d["Eyi"], d["Ezr"], d["Ezi"])
    else:
        return ret

def zGetPolTraceDirectArray(numRays, x=None, y=None, z=None, l=None, m=None,
                            n=None, Exr=None, Exi=None, Eyr=None, Eyi=None,
                            Ezr=None, Ezi=None, Ex=0, Ey=0, Phax=0, Phay=0,
                            intensity=None, waveNum=None, mode=0, startSurf=0,
                            lastSurf=-1, timeout=5000):
    """Trace large number of polarized rays defined by the ``x``, ``y``, ``z``, 
    ``l``, ``m`` and ``n`` coordinates on any starting surface as well as electric 
    field magnitude and relative phase. Similar to ``GetPolTraceDirect()``

    Ray tracing is performed on the lens file in the LDE of main Zemax 
    application (not in the DDE server)

    Parameters
    ----------
    numRays : integer
        number of rays to trace. ``numRays`` should be equal to the length
        of the lists (if provided) ``hx``, ``hy``, ``px``, etc.
    x : list, optional
        list specifying the x coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``x`` is created.
    y : list, optional
        list specifying the y coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``y`` is created
    z : list, optional
        list specifying the z coordinates of the ray at the start surface,
        of length ``numRays``; if ``None``, a list of 0.0s for ``z`` is created.
    l : list, optional
        list of x-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``l`` is created
    m : list, optional
        list of y-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``m`` is created
    n : list, optional
        list of z-direction cosines, of length ``numRays``; if ``None``, a
        list of 0.0s for ``n`` is created
    Exr : list, optional
        list of real part of the electric field in x direction for each ray.
        if ``None``, a list of 0.0s for ``Exr`` is created. See Notes
    Exi : list, optional
        list of imaginary part of the electric field in x direction for each ray.
        if ``None``, a list of 0.0s for ``Exi`` is created. See Notes
    Eyr : list, optional
        list of real part of the electric field in y direction for each ray.
        if ``None``, a list of 0.0s for ``Eyr`` is created. See Notes
    Eyi : list, optional
        list of imaginary part of the electric field in y direction for each ray.
        if ``None``, a list of 0.0s for ``Eyi`` is created. See Notes
    Ezr : list, optional
        list of real part of the electric field in z direction for each ray.
        if ``None``, a list of 0.0s for ``Ezr`` is created. See Notes
    Ezi : list, optional
        list of imaginary part of the electric field in z direction for each ray.
        if ``None``, a list of 0.0s for ``Ezi`` is created. See Notes
    Ex : float
        normalized electric field magnitude in x direction to be defined in
        array position 0. If not provided, an unpolarized ray will be traced.
        See Notes.
    Ey : float
        normalized electric field magnitude in y direction to be defined in
        array position 0. If not provided, an unpolarized ray will be traced.
        See Notes.
    Phax : float
        relative phase in x direction in degrees
    Phay : float
        relative phase in y direction in degrees
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
    startSurf : integer, optional
        start surface number (default = 0)
    lastSurf : integer, optional
        surface to trace the ray to, which is any valid surface number
        (``surf = -1``, default)
    timeout : integer, optional
        command timeout specified in milli-seconds

    Returns
    -------
    error : list of integers
        0 = ray traced successfully;
        +ve number = the ray missed the surface;
        -ve number = the ray total internal reflected (TIR) at surface
                     given by the absolute value of the ``error``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.
    Exr : list of real values
        list of real parts of the electric field components in x
    Exi : list of real values
        list of imaginary parts of the electric field components in x
    Eyr : list of real values
        list of real parts of the electric field components in y
    Eyi : list of real values
        list of imaginary parts of the electric field components in y
    Ezr : list of real values
        list of real parts of the electric field components in z
    Ezi : list of real values
        list of imaginary parts of the electric field components in z

    If ray tracing fails, a single integer error code is returned,
    which has the following meaning: -1 = Couldn't retrieve data in
    PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
    -998 = timeout reached

    Notes
    -----
    1. If all six of the electric field values Exr, Exi, Eyr, Eyi, Ezr, and Ezi
       for a ray are zero; then Zemax will use the ``Ex`` and ``Ey`` values
       provided in array position 0 to determine the electric field. Otherwise,
       the electric field is defined by these six values.
    2. The defined electric field vector must be orthogonal to the ray vector
       or incorrect ray tracing will result.
    3. Even if these six values are defined for each ray, values for ``Ex`` and
       ``Ey`` in the array position 0 must still be defined, otherwise an
       unpolarized ray trace will result.
    """
    rd = getRayDataArray(numRays, tType=3, mode=mode, startSurf=startSurf,
                         endSurf=lastSurf, x=Ex, y=Ey, z=Phax, l=Phay)
    x = x if x else [0.0] * numRays
    y = y if y else [0.0] * numRays
    z = z if z else [0.0] * numRays
    l = l if l else [0.0] * numRays
    m = m if m else [0.0] * numRays
    n = n if n else [0.0] * numRays
    Exr = Exr if Exr else [0.0] * numRays
    Exi = Exi if Exi else [0.0] * numRays
    Eyr = Eyr if Eyr else [0.0] * numRays
    Eyi = Eyi if Eyi else [0.0] * numRays
    Ezr = Ezr if Ezr else [0.0] * numRays
    Ezi = Ezi if Ezi else [0.0] * numRays

    if intensity:
        intensity = intensity if isinstance(intensity, list) else [intensity]*numRays
    else:
        intensity = [1.0] * numRays
    if waveNum:
        waveNum = waveNum if isinstance(waveNum, list) else [waveNum]*numRays
    else:
        waveNum = [1] * numRays

    # fill up the structure
    for i in xrange(1, numRays+1):
        rd[i].x = x[i-1]
        rd[i].y = y[i-1]
        rd[i].z = z[i-1]
        rd[i].l = l[i-1]
        rd[i].m = m[i-1]
        rd[i].n = n[i-1]
        rd[i].intensity = intensity[i-1]
        rd[i].Exr = Exr[i-1]
        rd[i].Exi = Exi[i-1]
        rd[i].Eyr = Eyr[i-1]
        rd[i].Eyi = Eyi[i-1]
        rd[i].Ezr = Ezr[i-1]
        rd[i].Ezi = Ezi[i-1]
        rd[i].wave = waveNum[i-1]
        # error & vigcode set to 0, opd and want_opd ignored

    # call ray tracing
    ret = zArrayTrace(rd, timeout)
    d = {}
    if ret == 0:
        reals = ['intensity', 'Exr', 'Exi', 'Eyr', 'Eyi', 'Ezr', 'Ezi']
        ints = ['error', ]
        for r in reals:
            exec(r + " = [0.0] * numRays", locals(), d)
        for i in ints:
            exec(i + " = [0] * numRays", locals(), d)
        for i in xrange(1, numRays+1):
            d["intensity"][i-1] = rd[i].intensity
            d["Exr"][i-1] = rd[i].Exr
            d["Exi"][i-1] = rd[i].Exi
            d["Eyr"][i-1] = rd[i].Eyr
            d["Eyi"][i-1] = rd[i].Eyi
            d["Ezr"][i-1] = rd[i].Ezr
            d["Ezi"][i-1] = rd[i].Ezi
            d["error"][i-1] = rd[i].error
        return (d["error"], d["intensity"], 
                d["Exr"], d["Exi"], d["Eyr"], d["Eyi"], d["Ezr"], d["Ezi"])
    else:
        return ret

def zGetNSCTraceArray(x=0.0, y=0.0, z=0.0, l=0.0, m=0.0, n=1.0, Exr=0.0, Exi=0.0,
                      Eyr=0.0, Eyi=0.0, Ezr=0.0, Ezi=0.0, intensity=1.0, waveNum=0,
                      surf=1, insideOf=0, usePolar=0, split=0, scatter=0, nMaxSegments=50,
                      timeout=5000):
    """Trace a single ray inside a non-sequential group. Rays may split or
    scatter into multiple paths. The function returns the entire tree of
    ray data containing split and/or scattered rays.
    
    Ray tracing is performed on the lens file in the LDE of main Zemax 
    application (not in the DDE server)

    Parameters
    ----------
    x : real
        starting x coordinate
    y : real
        starting y coordinate
    z : real
        starting z coordinate
    l : real
        starting x direction cosine
    m : real
        starting y direction cosine
    n : real
        starting z direction cosine
    Exr, Exi, Eyr, Eyi, Ezr, Ezi : reals
        initial real and imaginary electric fields along x, y, and z. They
        are required if performing polarization ray tracing
    intensity : real
        initial intensity
    waveNum : integer
        wavelength number, use 0 for randomly selected by weight
    surf : integer
        NSC group surface, 1 if the program mode is NSC
    insideOf : integer
        indicates where the ray starts. use 0 if ray is not inside anything
    usePolar : integer (0 or 1)
        0 = not performing polarization ray tracing, 1 = polarization ray
        tracing
    split : integer (0 or 1)
        0 = ray splitting is not used, 1 = ray splitting used
    scatter : integer (0 or 1)
        0 = scattering is not used, 1 = ray scattering used
    nMaxSegments : integer
        maximum allowed size of the raydata array for Zemax to return. The
        nMaxSegments value Zemax is using for the current optical system
        may be determined by the ``zGetNSCSettings()`` command.

    Returns
    -------
    rayData : list
        list of ray segments contaning the entire ray tree. Each element of the
        list is a ray segment (a Python named tuple), that can be retrieved as::

            totalSegments = len(rayData)
            for seg in rayData:
                segment_level = seg.segment_level
                segment_parent = seg.segment_parent
                inside_of_object_number = seg.inside_of
                hit_object_number = seg.hit_object
                x, y, z, l, m, n = seg.x, seg.y, seg.z, seg.l, seg.m, seg.n
                intensity = seg.intensity
                optical_path_length = seg.opl  # optical path length to hit object

    If ray tracing fails, a single integer error code is returned,
    which has the following meaning: -1 = Couldn't retrieve data in
    PostArrayTraceMessage, -999 = Couldn't communicate with Zemax,
    -998 = timeout reached

    Notes
    -----
    1. If polarization ray tracing is used, the initial electric field values
       must be provided. The user application must ensure the defined vector
       is orthogonal to the ray propagation vector, and that the resulting
       intensity matches the starting intensity value, otherwise incorrect
       ray tracing results will be produced.
    2. If ray splitting is to be used, polarization must be used as well.
    """
    assert not(split and not usePolar), "Polarization must be used, if splitting is used"
    vigcode = int(usePolar + 2*split + 4*scatter)
    assert 0 <= vigcode <= 7, "Total of pol + split + scatter must be in [0, 7]"
    # create ray data array of nMaxSegments + 10 elements (the "10" is really
    # quite arbitrary)
    rd = getRayDataArray(numRays=nMaxSegments+10, x=x, y=y, z=z, l=l, m=m,
                         n=n, opd=nMaxSegments+5, intensity=intensity,
                         Exr=Exr, Exi=Exi, Eyr=Eyr, Eyi=Eyi, Ezr=Ezr, Ezi=Ezi,
                         wave=waveNum, error=surf, vigcode=vigcode, want_opd=insideOf)
    # call ray tracing
    ret = zArrayTrace(rd, timeout)
    # parse the ray tree
    if ret == 0:
        segData = _co.namedtuple('segment', ['segment_level', 'segment_parent',
                                             'inside_of', 'hit_object', 'x',
                                             'y', 'z', 'l', 'm', 'n', 'intensity',
                                             'opl'])
        nNumRaySegments = rd[0].want_opd # total number of segments stored
        rayData = ['']*nNumRaySegments
        for i in xrange(1, nNumRaySegments+1):
            rayData[i-1] = segData(rd[i].wave, rd[i].want_opd, rd[i].vigcode,
                                   rd[i].error, rd[i].x, rd[i].y, rd[i].z,
                                   rd[i].l, rd[i].m, rd[i].n, rd[i].intensity, rd[i].opd)
        return rayData
    else:
        return ret

# ###########################################################################
# Basic test functions
# Please note that some of the following tests require Zemax to be running
# In addition, load a lens file into Zemax
# ###########################################################################

def _test_getRayDataArray():
    """test the getRayDataArray() function
    """
    print("Basic test of getRayDataArray() function:")
    # create RayData without any kwargs
    rd = getRayDataArray(numRays=5)
    assert len(rd) == 6
    assert rd[0].error == 5     # number of rays
    assert rd[0].opd == 0       # GetTrace ray tracing type
    assert rd[0].wave == 0      # real ray tracing
    assert rd[0].want_opd == -1  # last surface

    # create RayData with some more arguments
    rd = getRayDataArray(numRays=5, tType=3, mode=1, startSurf=2)
    assert rd[0].opd == 3        # mode 3
    assert rd[0].wave == 1       # real ray tracing
    assert rd[0].vigcode == 2    # first surface

    # create RayData with kwargs
    rd = getRayDataArray(numRays=5, tType=2, x=1.0, y=1.0)
    assert rd[0].x == 1.0
    assert rd[0].y == 1.0
    assert rd[0].z == 0.0

    # create RayData with kwargs overriding some regular parameters
    rd = getRayDataArray(numRays=5, tType=2, x=1.0, y=1.0, error=1)
    assert len(rd) == 6
    assert rd[0].x == 1.0
    assert rd[0].y == 1.0
    assert rd[0].z == 0.0
    assert rd[0].error == 1

def _test_arraytrace_module_basic():
    """very basic test for the arraytrace module
    """
    # Basic test of the module functions
    print("Basic test of arraytrace module:")
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
    print("OPDs:")
    for i in range(1, 11):  # SEEMS LIKE ZEMAX IS CALCULATING THE OPD EVEN WHEN NOT ASKED
        print(rd[k].opd)
    print("Intensities:")
    for i in range(1, 11):
        print(rd[k].intensity)
    print("Success!")

if __name__ == '__main__':
    # run the test functions
    _test_getRayDataArray()
    _test_arraytrace_module_basic()