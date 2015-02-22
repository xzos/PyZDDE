#-------------------------------------------------------------------------------
# Name:        arraytrace.py
# Purpose:     Module for doing array ray tracing in Zemax.
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2015
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.8.02
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

def getRayDataArray(numRays, tType=0, mode=0, startSurf=None, endSurf=-1,
                    **kwargs):
    """function to create the basic ray data structure array, ``rd``.

    The caller must fill the rest of the appropriate fields of the
    array from elements ``rd[1]`` to ``rd[numRays]``

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
    kwargs : keyword arguments
        The ``kwargs`` are used to specify additional fields of the 0th ray
        data array element such as ``x``, ``y``, ``z``, and ``l`` for ``Ex``,
        ``Ey``, ``Phax``, and ``Phay`` in ``zGetPolTraceArray()`` and in
        ``zGetPolTraceDirectArray()`` (``tType=2``). Please use the generic
        field name of the ``DdeArrayData`` as argument names.

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
    other_fields = {'x', 'y', 'z', 'l', 'm', 'n',
                    'Exr', 'Exi', 'Eyr', 'Eyi', 'Ezr', 'Ezi'}
    if kwargs:
        assert set(kwargs).issubset(other_fields), "Received one or more unexpected kwargs "
    # create ctypes array
    rd = (DdeArrayData * (numRays + 1))()
    # Setup a basic ray data array for test
    rd[0].opd = _ct.c_double(tType)
    rd[0].wave = _ct.c_int(mode)
    rd[0].error = _ct.c_int(numRays)
    if startSurf:
        rd[0].vigcode = _ct.c_int(startSurf)
    rd[0].want_opd = _ct.c_int(endSurf)
    # fill up based on kwargs
    if kwargs:
        for k in kwargs:
            setattr(rd[0], k, kwargs[k])
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
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the x, y, and z direction cosines after refraction into the media
        following the requested surface.
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


def zGetTraceDirectArray(numRays, x=None, y=None, z=None, l=None, m=None,
                         n=None, intensity=None, waveNum=None, mode=0,
                         startSurf=0, lastSurf=-1, timeout=5000):
    """Trace large number of rays defined by ``x``, ``y``, ``z``, ``l``,
    ``m`` and ``n`` coordinates on any starting surface as well as
    wavelength number, mode and the surface to trace the ray to.

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
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the x, y, and z direction cosines after refraction into the media
        following the requested surface.
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
        print("Error. zArrayTraceDirect returned error code {}".format(ret))

def zGetPolTraceArray(numRays, hx=None, hy=None, px=None, py=None, Exr=None,
                      Exi=None, Eyr=None, Eyi=None, Ezr=None, Ezi=None, Ex=0,
                      Ey=0, Phax=0, Phay=0, intensity=None, waveNum=None, mode=0,
                      surf=-1, timeout=5000):
    """Trace large number of polarized rays defined by their normalized
    field and pupil coordinates. Similar to ``GetPolTrace()``

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
    x, y, z : list of reals
        x, or , y, or z, coordinates of the ray on the requested surface
    l, m, n : list of reals
        the x, y, and z direction cosines after refraction into the media
        following the requested surface.
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
    1. If all six of the electric field values Exr, Exi, Eyr, Eyi, Ezr, and Ezi
       for a ray are zero; then ZEMAX will use the ``Ex`` and ``Ey`` values
       provided in array position 0 to determine the electric field. Otherwise,
       the electric field is defined by these six values.
    2. The defined electric field vector must be orthogonal to the ray vector
       or incorrect ray tracing will result.
    3. Even if these six values are defined for each ray, values for ``Ex`` and
       ``Ey`` in the array position 0 must still be defined, otherwise an
       unpolarized ray trace will result.
    """
    rd = getRayDataArray(numRays, tType=2, mode=mode, endSurf=surf)
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

def _test_zGetArrayTrace_basic():
    """test zGetArrayTrace() function
    """
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

def _test_zGetPolTraceArray():
    """test zGetPolTraceArray() function
    """
    # TO DO
    pass


if __name__ == '__main__':
    # run the test functions
    _test_getRayDataArray()
    _test_arraytrace_module_basic()
    _test_zGetArrayTrace_basic()