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
    2. zGetOPDArray()
    ToDo: 2. zGetTraceDirectArray()
    ToDo: 3. zGetPolTraceArray()
    ToDo: 4. zGetPolTraceDirectArray()
    ToDo: 5. zGetNSCTraceArray()
    
Note
-----

The parameter want_opd is very confusing as it alters the behavior of GetTraceArray.
If the calculation of OPD is requested for a ray, 
- vigcode becomes a different meaning (seems to be 1 if vignetted, but no longer related to surface)
- bParaxial (mode) becomes inactive, Zemax always performs a real-raytrace !
- the surface normal is not calculated
- if the calculation of the chief ray data is not requested for the first ray, 
  e.g. by setting all want_opd to 1, wrong OPD values are returned (without any relation to the real values)
- this affects only rays with want_opd<>0 -> i.e. if it is mixed, one obtains a total mess


The pupil apodization seems to be always considered independent of the mode / bParaxial value
in contrast to the note in the Zemax Manual (tested with Zemax 13 Release 2 SP 5 Premium 64bit)
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
#  Make sure the arrays are C-contigous, see
#http://stackoverflow.com/questions/26998223/what-is-the-difference-between-contiguous-and-non-contiguous-arrays
#  Make sure the input array is aligned on proper boundaries for its data type.                         
_INT  = _ct.c_int;
_INT1D = _np.ctypeslib.ndpointer(ndim=1,dtype=_np.int,flags=["C_CONTIGUOUS","ALIGNED"])
_INT2D = _np.ctypeslib.ndpointer(ndim=2,dtype=_np.int,flags=["C_CONTIGUOUS","ALIGNED"])
_DBL1D = _np.ctypeslib.ndpointer(ndim=1,dtype=_np.double,flags=["C_CONTIGUOUS","ALIGNED"])
_DBL2D = _np.ctypeslib.ndpointer(ndim=2,dtype=_np.double,flags=["C_CONTIGUOUS","ALIGNED"])
  
def _init_list1d(var,length,dtype,name):
  "repeat variable if it is scalar to become a 1d list of size length"
  var=_np.atleast_1d(var);  
  if var.size==1: 
    return _np.full(length,var[0],dtype=dtype);
  else:     
    assert var.size==length, 'incompatible length of list \'%s\': should be %d instead of %d'%(name,length,var.size)
    return var.flatten();
    
    
def zGetTraceArray(hx,hy, px,py, intensity=1., waveNum=1,
                   bParaxial=False, surf=-1, timeout=60000):
    """Trace large number of rays defined by their normalized field and pupil
    coordinates on lens file in the LDE of main Zemax application (not in the DDE server)

    Parameters
    ----------
    hx,hy : float or vector of length ``numRays``
        list of normalized field heights along x and y axis
    px,py : float or vector of length ``numRays``
        list of normalized heights in pupil coordinates, along x and y axis
    intensity : float or vector of length ``numRays``, optional
        initial intensities. If a vector of length ``numRays`` is given it is
        used. If a single float value is passed, all rays use the same value for
        their initial intensities. Default: all intensities are set to ``1.0``.
    waveNum : integer or vector of length ``numRays``, optional
        wavelength number. If a vector of integers of length ``numRays`` is given
        it is used. If a single integer value is passed, all rays use the same
        value for wavelength number. Default: wavelength number equal to 1.
    bParaxial : bool, optional
        If True, a paraxial raytrace is performed (default: False, real raytrace)
    surf : integer, optional
        surface to trace the ray to. Usually, the ray data is only needed at
        the image surface (``surf = -1``, default)
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
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined. 

    If ray tracing fails, an RuntimeError is raised.
    
    Examples
    -------- 
    >>> import numpy as np     
    >>> import matplotlib.pylab as plt
    >>> # cartesian sampling in field an pupil
    >>> x = np.linspace(-1,1,4)
    >>> p = np.linspace(-1,1,3)    
    >>> hx,hy,px,py = np.meshgrid(x,x,p,p);
    >>> # run array-trace
    >>> (error,vigcode,pos,dir,normal,intensity) = \\
          zGetTraceArray(hx,hy,px,py);
    >>> # plot results
    >>> plt.scatter(pos[:,0],pos[:,1],c=_np.sqrt(hx**2+hy**2))
    """
    
    # handle input arguments 
    hx,hy,px,py = _np.atleast_1d(hx,hy,px,py);
    nRays = max(hx.size,hy.size,px.size,py.size);
    hx = _init_list1d(hx,nRays,_np.double,'hx');
    hy = _init_list1d(hy,nRays,_np.double,'hy');
    px = _init_list1d(px,nRays,_np.double,'px');
    py = _init_list1d(py,nRays,_np.double,'py');
    intensity = _init_list1d(intensity,nRays,_np.double,'intensity');
    waveNum   = _init_list1d(waveNum,nRays,_np.int,'waveNum');
    
    # allocate memory for return values
    error=_np.zeros(nRays,dtype=_np.int);                   
    vigcode=_np.zeros(nRays,dtype=_np.int);
    pos=_np.zeros((nRays,3),dtype=_np.double);
    dir=_np.zeros((nRays,3),dtype=_np.double);
    normal=_np.zeros((nRays,3),dtype=_np.double);
                   
    # numpyGetTrace(int nrays, double hx[], double hy[], double px[], double py[],
    #   double intensity[], int wave_num[], int mode, int surf, int error[], 
    #   int vigcode[], double pos[][3], double dir[][3], double normal[][3], unsigned int timeout);
    _numpyGetTrace = _array_trace_lib.numpyGetTrace
    _numpyGetTrace.restype = _INT
    _numpyGetTrace.argtypes= [_INT,_DBL1D,_DBL1D,_DBL1D,_DBL1D,_DBL1D,_INT1D,_INT,
                              _INT,_INT1D,_INT1D,_DBL2D,_DBL2D,_DBL2D,_ct.c_uint]
    ret = _numpyGetTrace(nRays,hx,hy,px,py,intensity,waveNum,int(bParaxial),
                         surf,error,vigcode,pos,dir,normal,timeout)
    # analyse error - flag
    if ret==-1: raise RuntimeError("Couldn't retrieve data in PostArrayTraceMessage.")
    if ret==-999: raise RuntimeError("Couldn't communicate with Zemax.");
    if ret==-998: raise RuntimeError("Timeout reached after %dms"%timeout);
  
    return (error,vigcode,pos,dir,normal,intensity);


def zGetOpticalPathDifferenceArray(hx,hy, px,py, waveNum=1,timeout=60000):
    """
    Calculates the optical path difference (OPD) between real ray and 
    real chief ray at the image plane for a large number of rays. The lens 
    file in the LDE of main Zemax application (not in the DDE server) is used.

    Parameters
    ----------
    hx,hy : float or vector of length ``nField``
        list of normalized field heights along x and y axis
    px,py : float or vector of length ``nPupil``
        list of normalized heights in pupil coordinates, along x and y axis
    waveNum : integer or vector of length ``nWave``, optional
        list of wavelength numbers. Default: single wavelength number equal to 1.
    timeout : integer, optional
        command timeout specified in milli-seconds (default: 1min), at least 1s

    Returns
    -------
    error : ndarray of shape (nWave,nField,nPupil)
        * =0: ray traced successfully
        * <0: ray missed the surface number indicated by ``error``
        * >0: total internal reflection of ray at the surface number given by ``-error``
    vigcode : ndarray of shape (nWave,nField,nPupil)
        indicates if ray was vignetted (vigcode==1) or not (vigcode=0)
    opd : ndarray of shape (nWave,nField,nPupil)
        computed optical path difference in waves, corresponding to the 
        wavelength of the individual ray
    pos : ndarray of shape (nWave,nField,nPupil,3)
        local image coordinates ``(x,y,z)`` of each ray
    dir : ndarray of shape (nWave,nField,nPupil,3)
        local direction cosines ``(l,m,n)`` at the image surface
    intensity : ndarray of shape (nWave,nField,nPupil)
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined.

    If ray tracing fails, an RuntimeError is raised.
    
    Examples
    -------- 
    >>> import numpy as np     
    >>> import matplotlib.pylab as plt
    >>> # pupil sampling along diagonal (px,px)
    >>> NP=51; p = _np.linspace(-1,1,NP);
    >>> (error,vigcode,opd,pos,dir,intensity) = \\
          zGetOpticalPathDifferenceArray(0,0,p,p);
    >>> plt.plot(p,opd[0,0,:])
    """

    # handle input arguments 
    hx,hy,px,py,waveNum = _np.atleast_1d(hx,hy,px,py,waveNum);
    nField = max(hx.size,hy.size)
    nPupil = max(px.size,py.size);
    nWave  = waveNum.size;
    nRays  = nWave*nField*nPupil;    
    hx = _init_list1d(hx,nField,_np.double,'hx');
    hy = _init_list1d(hy,nField,_np.double,'hy');
    px = _init_list1d(px,nPupil,_np.double,'px');
    py = _init_list1d(py,nPupil,_np.double,'py');
        
    # allocate memory for return values
    error=_np.zeros(nRays,dtype=_np.int);                   
    vigcode=_np.zeros(nRays,dtype=_np.int);
    opd=_np.zeros(nRays,dtype=_np.double);
    pos=_np.zeros((nRays,3),dtype=_np.double);
    dir=_np.zeros((nRays,3),dtype=_np.double);
    intensity=_np.zeros(nRays,dtype=_np.double);
                   
    # int numpyOpticalPathDifference(int nField, double hx[], double hy[],
    #          int nPupil, double px[], double py[], int nWave, int wave_num[], 
    #          int error[], int vigcode[], double opd[], double pos[][3], 
    #          double dir[][3], double intensity[], unsigned int timeout)
    # result arrays are multidimensional arrays indexed as [wave][field][pupil]
    _numpyOPD = _array_trace_lib.numpyOpticalPathDifference
    _numpyOPD.restype = _INT
    _numpyOPD.argtypes= [_INT,_DBL1D,_DBL1D, _INT,_DBL1D,_DBL1D, _INT,_INT1D,
                              _INT1D,_INT1D,_DBL1D,_DBL2D,_DBL2D,_DBL1D,_ct.c_uint]
    ret = _numpyOPD(nField,hx,hy,nPupil,px,py,nWave,waveNum,
                   error,vigcode,opd,pos,dir,intensity,timeout)
    # analyse error - flag
    if ret==-1: raise RuntimeError("Couldn't retrieve data in PostArrayTraceMessage.")
    if ret==-999: raise RuntimeError("Couldn't communicate with Zemax.");
    if ret==-998: raise RuntimeError("Timeout reached after %dms"%timeout);
    
    # reshape arrays as multidimensional arrays
    s1=(nWave,nField,nPupil);
    s3=(nWave,nField,nPupil,3);
  
    return (error.reshape(s1),vigcode.reshape(s1),opd.reshape(s1),
            pos.reshape(s3),dir.reshape(s3),intensity.reshape(s1));


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
    x = _np.linspace(-1,1,4)
    p = _np.linspace(-1,1,3)    
    hx,hy,px,py = _np.meshgrid(x,x,p,p);
    (error,vigcode,pos,dir,normal,intensity) = \
        zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=-1);
    
    print(" number of rays: %d" % len(pos));
    if len(pos)<1e5:
      import matplotlib.pylab as plt
      from mpl_toolkits.mplot3d import Axes3D
      fig = plt.figure()
      ax = fig.add_subplot(111,projection='3d')
      ax.scatter(*pos.T,c=_np.sqrt(px**2+py**2));
    print("Success!")


def _test_zOPDArray():
    """very basic test for the zGetTraceNumpy function
    """
    # Basic test of the module functions
    print("Basic test of zGetTraceNumpy module:")
    NP=21; px = _np.linspace(-1,1,NP); 
    NF=5;  hx = _np.linspace(-1,1,NF);
    (error,vigcode,opd,pos,dir,intensity) = \
        zGetOpticalPathDifferenceArray(hx,0,px,0);
    
    print(" number of rays: %d" % opd.size);
    if opd.size<1e5:
      import matplotlib.pylab as plt
      plt.figure();
      for f in xrange(NF):
        plt.plot(px,opd[0,f],label="hx=%5.3f"%hx[f]);
      plt.legend(loc=0);
    print("Success!")


if __name__ == '__main__':
    # run the test functions
    _test_zGetTraceArray()
    _test_zOPDArray()
