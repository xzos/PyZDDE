# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        numpy_interface.py
# Purpose:     Module for doing array ray tracing in Zemax using numpy arrays 
#              to pass ray data to Zemax
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
"""
Module for doing array ray tracing as described in Zemax manual. The following
functions are provided according to 5 different modes discussed in the Zemax manual:

    1. zGetTraceArray()
    2. zGetTraceDirectArray()
    ToDo: 3. zGetPolTraceArray()
    ToDo: 4. zGetPolTraceDirectArray()
    ToDo: 5. zGetNSCTraceArray()
"""
import os as _os
import sys as _sys
import ctypes as _ct
import numpy as _np

if _sys.version_info[0] > 2:
    xrange = range

def _is64bit():
    """return True if Python version is 64 bit
    """
    return _sys.maxsize > 2**31 - 1

# load the arrayTrace library
_dllDir = "x64\\Release\\" if _is64bit() else "win32\\Release\\"
_dllName = "ArrayTrace.dll"
_dllpath = _os.path.join(_os.path.dirname(_os.path.realpath(__file__)), _dllDir)
_array_trace_lib = _ct.WinDLL(_dllpath + _dllName)
                          
# shorthands for CTypes
_INT  = _ct.c_int;
_INT1D = _np.ctypeslib.ndpointer(ndim=1,dtype=_np.int,flags=["C_CONTIGUOUS","ALIGNED"])
_INT2D = _np.ctypeslib.ndpointer(ndim=2,dtype=_np.int,flags=["C_CONTIGUOUS","ALIGNED"])
_DBL1D = _np.ctypeslib.ndpointer(ndim=1,dtype=_np.double,flags=["C_CONTIGUOUS","ALIGNED"])
_DBL2D = _np.ctypeslib.ndpointer(ndim=2,dtype=_np.double,flags=["C_CONTIGUOUS","ALIGNED"])
  
def zGetTraceArray(field, pupil, intensity=None, waveNum=None,
                   mode=0, surf=-1, want_opd=0, timeout=60000):
    """Trace large number of rays defined by their normalized field and pupil
    coordinates on lens file in the LDE of main Zemax application (not in the DDE server)

    Parameters
    ----------
    field : ndarray of shape (``numRays``,2)
        list of normalized field heights along x and y axis
    px : ndarray of shape (``numRays``,2)
        list of normalized heights in pupil coordinates, along x and y axis
    intensity : float or vector of length ``numRays``, optional
        initial intensities. If a vector of length ``numRays`` is given it is
        used. If a single float value is passed, all rays use the same value for
        their initial intensities. Default: all intensities are set to ``1.0``.
    waveNum : integer or vector of length ``numRays``, optional
        wavelength number. If a vector of integers of length ``numRays`` is given
        it is used. If a single integer value is passed, all rays use the same
        value for wavelength number. Default: wavelength number equal to 1.
    mode : integer, optional
        0 = real (Default), 1 = paraxial
    surf : integer, optional
        surface to trace the ray to. Usually, the ray data is only needed at
        the image surface (``surf = -1``, default)
    want_opd : integer, optional
        0 if OPD data is not needed (Default), 1 if it is. See Zemax manual
        for details.
    timeout : integer, optional
        command timeout specified in milli-seconds (default: 1min), at least 1s

    Returns
    -------
    error : list of integers
        * =0: ray traced successfully
        * <0: ray missed the surface number indicated by ``error``
        * >0: total internal reflection of ray at the surface number given by ``-error``
    vigcode : list of integers
        the first surface where the ray was vignetted. Unless an error occurs
        at that surface or subsequent to that surface, the ray will continue
        to trace to the requested surface.
    pos : ndarray of shape (``numRays``,3)
        local coordinates ``(x,y,z)`` of each ray on the requested surface
    dir : ndarray of shape (``numRays``,3)
        local direction cosines ``(l,m,n)`` after refraction into the media
        following the requested surface.
    normal : ndarray of shape (``numRays``,3)
        local direction cosines ``(l2,m2,n2)`` of the surface normals at the 
        intersection point of the ray with the requested surface
    opd : list of reals
        computed optical path difference if ``want_opd <> 0``
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined. Note mode 0 considers pupil apodization, 
        while mode 1 does not.

    If ray tracing fails, an RuntimeError is raised.
    
    Examples
    -------- 
    >>> import numpy as np     
    >>> import matplotlib.pylab as plt
    >>> # cartesian sampling in field an pupil
    >>> x = np.linspace(-1,1,10)
    >>> px= np.linspace(-1,1,3)    
    >>> grid = np.meshgrid(x,x,px,px);
    >>> field= np.transpose(grid[0:2]).reshape(-1,2);
    >>> pupil= np.transpose(grid[2:4]).reshape(-1,2);
    >>> # run array-trace
    >>> (error,vigcode,pos,dir,normal,opd,intensity) = \\
    >>>      zGetTraceNumpy(field,pupil,mode=0);
    >>> # plot results
    >>> plt.scatter(pos[:,0],pos[:,1])

    Notes
    -----
    The opd can only be computed if the last surface is the image surface,
    otherwise, the opd value will be zero. It is not yet clear, if want_opd
    works as described in the manual: 
    
      If want_opd is less than zero(such as -1) then the both the chief ray and
      specified  ray  are requested, and the  OPD  is  the phase  difference  
      between  the two  in  waves  of  the current wavelength. If want_opd is
      greater than zero, then the most recently traced chief ray data is used. 
      Therefore, the want_opd flag should be -1 whenever the chief ray changes; 
      and +1 for all subsequent rays which do not require the chief ray be 
      traced again. Generally the chief ray changes only when the field 
      coordinates or wavelength changes.
    """
    # handle input arguments 
    assert 2 == field.ndim == pupil.ndim, 'field and pupil should be 2d arrays'                 
    assert field.shape == pupil.shape, 'we expect field and pupil points for each ray'
    nRays = field.shape[0];
    if intensity is None: intensity=1;
    if _np.isscalar(intensity): intensity=_np.zeros(nRays)+intensity;        
    if waveNum is None: waveNum=1;
    if _np.isscalar(waveNum): waveNum=_np.zeros(nRays,dtype=_np.int)+waveNum;
             
                     
    # set up output arguments
    error=_np.zeros(nRays,dtype=_np.int);                   
    vigcode=_np.zeros(nRays,dtype=_np.int);
    pos=_np.zeros((nRays,3));
    dir=_np.zeros((nRays,3));
    normal=_np.zeros((nRays,3));
    opd=_np.zeros(nRays);
                   
    # numpyGetTrace(int nrays, double field[][2], double pupil[][2], 
    #   double intensity[], int wave_num[], int mode, int surf, int want_opd, 
    #   int error[], int vigcode[], double pos[][3], double dir[][3], double normal[][3], 
    #  double opd[], unsigned int timeout);
    _numpyGetTrace = _array_trace_lib.numpyGetTrace
    _numpyGetTrace.restype = _INT
    _numpyGetTrace.argtypes= [_INT,_DBL2D,_DBL2D,_DBL1D,_INT1D,_INT,_INT,_INT,
                              _INT1D,_INT1D,_DBL2D,_DBL2D,_DBL2D,_DBL1D,_ct.c_uint]
    ret = _numpyGetTrace(nRays,field,pupil,intensity,waveNum,mode,surf,want_opd,
                   error,vigcode,pos,dir,normal,opd,timeout)
    # analyse error - flag
    if ret==-1: raise RuntimeError("Couldn't retrieve data in PostArrayTraceMessage.")
    if ret==-999: raise RuntimeError("Couldn't communicate with Zemax.");
    if ret==-998: raise RuntimeError("Timeout reached after %dms"%timeout);
  
    return (error,vigcode,pos,dir,normal,opd,intensity);



# ###########################################################################
# Basic test functions
# Please note that some of the following tests require Zemax to be running
# In addition, load a lens file into Zemax
# ###########################################################################

def _test_zGetTraceArray():
    """very basic test for the zGetTraceNumpy function
    """
    # Basic test of the module functions
    print("Basic test of zGetTraceNumpy module:")
    x = _np.linspace(-1,1,10)
    px= _np.linspace(-1,1,3)    
    grid = _np.meshgrid(x,x,px,px);
    field= _np.transpose(grid[0:2]).reshape(-1,2);
    pupil= _np.transpose(grid[2:4]).reshape(-1,2);
    
    (error,vigcode,pos,dir,normal,opd,intensity) = \
        zGetTraceArray(field,pupil,mode=0);
    
    print(" number of rays: %d" % len(pos));
    if len(pos)<1e5:
      import matplotlib.pylab as plt
      from mpl_toolkits.mplot3d import Axes3D
      fig = plt.figure()
      ax = fig.add_subplot(111, projection='3d')
      ax.scatter(*pos.T,c=opd);#_np.linalg.norm(pupil,axis=1));
    
    print("Success!")




if __name__ == '__main__':
    # run the test functions
    _test_zGetTraceArray()
    