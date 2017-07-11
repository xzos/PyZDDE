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
    3. zGetTraceDirectArray()
    ToDo: 3. zGetPolTraceArray()
    ToDo: 4. zGetPolTraceDirectArray()
    ToDo: 5. zGetNSCTraceArray()
    
Note
-----
The pupil apodization seems to be always considered independent of the mode / bParaxial value
in contrast to the note in the Zemax Manual (tested with Zemax 13 Release 2 SP 5 Premium 64bit)
"""
import os as _os
import sys as _sys
import ctypes as _ct
import numpy as _np

if _sys.version_info[0] < 3:
    range = xrange

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
    >>> import pyzdde.zdde as pyz
    >>> import pyzdde.arraytrace.numpy_interface as nt
    >>> ln = pyz.createLink()
    >>> # cartesian sampling in field an pupil
    >>> x = np.linspace(-1,1,4)
    >>> p = np.linspace(-1,1,3)    
    >>> hx,hy,px,py = np.meshgrid(x,x,p,p);
    >>> # run array-trace
    >>> (error,vigcode,pos,dir,normal,intensity) = nt.zGetTraceArray(hx,hy,px,py);
    >>> # plot results
    >>> plt.scatter(pos[:,0],pos[:,1],c=np.sqrt(hx**2+hy**2))
    >>> ln.close();
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


def zGetTraceDirectArray(startpos, startdir, intensity=1., waveNum=1,
                   bParaxial=False, startSurf=0, lastSurf=-1, timeout=60000):
    """
    Trace large number of rays defined by their initial position and direction from
    one surface to another for the lens file given in the LDE of the main Zemax 
    application (not in the DDE server).

    Parameters
    ----------
    startpos : ndarray of shape (``numRays``,3)
        starting point of each ray at the initial surface, given in local coordinates x,y,z
    startdir : ndarray of shape (``numRays``,3)
        starting direction of each reay at the initial surface, given as
        local direction cosines kx,ky,kz
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
    startSurf : integer, optional
        start surface number (default: 0)
    lastSurf : integer, optional
        surface to trace the ray to. (default: -1, image surface)
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
        local coordinates ``(x,y,z)`` of each ray on the requested last surface
    dir : ndarray of shape (``numRays``,3)
        local direction cosines ``(l,m,n)`` after refraction into the media
        following the requested last surface.
    normal : ndarray of shape (``numRays``,3)
        local direction cosines ``(l2,m2,n2)`` of the surface normals at the 
        intersection point of the ray with the requested last surface
    intensity : list of reals
        the relative transmitted intensity of the ray, including any pupil
        or surface apodization defined. 

    If ray tracing fails, an RuntimeError is raised.
    
    Examples
    -------- 
    >>> import numpy as np
    >>> import matplotlib.pylab as plt
    >>> import pyzdde.zdde as pyz
    >>> import pyzdde.arraytrace.numpy_interface as nt
    >>> ln = pyz.createLink()
    >>> # launch rays from same from off-axis field point
    >>> # we create initial pos and dir using zGetTraceArray
    >>> nRays=7;  
    >>> startsurf= 1;  # in case of collimated input beam    
    >>> lastsurf = ln.zGetNumSurf();
    >>> hx,hy,px,py = 0, 0.5, 0, np.linspace(-1,1,nRays);
    >>> (_,_,pos,dir,_,_) = nt.zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=startsurf);    
    >>> # trace ray until last surface
    >>> points = np.zeros((lastsurf+1,nRays,3));    # indexing: surf,ray,coord
    >>> z0=0; points[startsurf]=pos;                # ray intersection points on starting surface
    >>> for isurf in range(startsurf,lastsurf):
    >>>   # trace to next surface
    >>>   (error,vigcode,pos,dir,_,_)=nt.zGetTraceDirectArray(pos,dir,bParaxial=False,startSurf=isurf,lastSurf=isurf+1);
    >>>   points[isurf+1]=pos;
    >>>   points[isurf+1,vigcode!=0]=np.nan;        # remove vignetted rays
    >>>   # add thickness of current surface (assumes absence of tilts or decenters in system)
    >>>   z0+=ln.zGetThickness(isurf);
    >>>   points[isurf+1,:,2]+=z0;
    >>>   print("  surface #%d at z-position z=%f" % (isurf+1,z0));
    >>> # plot rays in y-z section
    >>> x,y,z = points[startsurf:].T;
    >>> ax=plt.subplot(111,aspect='equal')
    >>> ax.plot(z.T,y.T,'.-')
    >>> ln.close();
    """
    # ensure correct memory alignment of input arrays
    startpos = _np.require(startpos,dtype=_np.double,requirements=['C','A']); 
    startdir = _np.require(startdir,dtype=_np.double,requirements=['C','A']);     
    intensity=_np.require(intensity,dtype=_np.double,requirements=['C','A']); 
    waveNum  =_np.require(waveNum,dtype=_np.int,requirements=['C','A']); 
     
    # handle input arguments 
    nRays = startpos.shape[0];
    assert (nRays,3) == startpos.shape == startdir.shape, 'startpos and startdir should have shape (nRays,3)'
    if intensity.size==1: intensity=_np.full(nRays,intensity,dtype=_np.double);        
    if waveNum.size==1: waveNum=_np.full(nRays,waveNum,dtype=_np.int);
    assert intensity.shape == (nRays,), 'intensity must be scalar or a vector of length nRays'        
    assert waveNum.shape == (nRays,), 'waveNum must be scalar or a vector of length nRays'
    
    # allocate memory for return values
    error=_np.zeros(nRays,dtype=_np.int);                   
    vigcode=_np.zeros(nRays,dtype=_np.int);
    pos=_np.zeros((nRays,3),dtype=_np.double);
    dir=_np.zeros((nRays,3),dtype=_np.double);
    normal=_np.zeros((nRays,3),dtype=_np.double);
                   
    # numpyGetTraceDirect(int nrays, double startpos[][3], double startdir[][3], 
    #   double intensity[], int wave_num[], int mode, int startsurf, int lastsurf, int error[], 
    #   int vigcode[], double pos[][3], double dir[][3], double normal[][3], unsigned int timeout)    
    _numpyGetTraceDirect = _array_trace_lib.numpyGetTraceDirect
    _numpyGetTraceDirect.restype = _INT
    _numpyGetTraceDirect.argtypes= [_INT,_DBL2D,_DBL2D,_DBL1D,_INT1D,_INT,_INT,_INT,
                                    _INT1D,_INT1D,_DBL2D,_DBL2D,_DBL2D,_ct.c_uint]
    ret = _numpyGetTraceDirect(nRays,startpos,startdir,intensity,waveNum,
            int(bParaxial),startSurf,lastSurf,error,vigcode,pos,dir,normal,timeout)
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
    >>> import pyzdde.zdde as pyz
    >>> import pyzdde.arraytrace.numpy_interface as nt
    >>> ln = pyz.createLink()
    >>> # pupil sampling along diagonal (px,px)
    >>> NP=51; p = np.linspace(-1,1,NP);
    >>> (error,vigcode,opd,pos,dir,intensity) = nt.zGetOpticalPathDifferenceArray(0,0,p,p);
    >>> plt.plot(p,opd[0,0,:])
    >>> ln.close();
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
    "very basic test for the zGetTraceNumpy function"
    # Basic test of the module functions
    print("Basic test of zGetTraceArray:")
    x = _np.linspace(-1,1,4)
    p = _np.linspace(-1,1,3)    
    hx,hy,px,py = _np.meshgrid(x,x,p,p);
    (error,vigcode,pos,dir,normal,intensity) = \
        zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=-1);
    
    print(" number of rays: %d" % len(pos));
    if len(pos)<1e5:
      import matplotlib.pylab as plt
      fig = plt.figure()
      ax = fig.add_subplot(111,projection='3d')
      ax.scatter(*pos.T,c=_np.sqrt(px**2+py**2));
    
    
def _test_zGetTraceArrayDirect():
    "very basic test for the zGetTraceNumpy function"    
    # Basic test of the module functions
    print("Basic test of zGetTraceArrayDirect:")
    x = _np.linspace(-1,1,4)
    p = _np.linspace(-1,1,3)    
    hx,hy,px,py = _np.meshgrid(x,x,p,p);
    # compare with zGetTraceArray results
    (_,_,startpos,startdir,_,_) = \
        zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=0);
    ref = zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=-1);
    # GetTraceArrayDirect
    ret = zGetTraceDirectArray(startpos,startdir,bParaxial=False,startSurf=0,lastSurf=-1);
    ret_descr = ('error','vigcode','pos','dir','normal','intensity')
    for i in range(len(ref)):
      assert _np.allclose(ref[i],ret[i]), "difference in %s"%ret_descr[i];

def _test_zOPDArray():
    "very basic test for the zGetTraceNumpy function"
    # Basic test of the module functions
    print("Basic test of zGetOpticalPathDifference:")
    NP=21; px = _np.linspace(-1,1,NP); 
    NF=5;  hx = _np.linspace(-1,1,NF);
    (error,vigcode,opd,pos,dir,intensity) = \
        zGetOpticalPathDifferenceArray(hx,0,px,0);
    
    print(" number of rays: %d" % opd.size);
    if opd.size<1e5:
      import matplotlib.pylab as plt
      plt.figure();
      for f in range(NF):
        plt.plot(px,opd[0,f],label="hx=%5.3f"%hx[f]);
      plt.legend(loc=0);

def _plot_2D_layout_from_array_trace():
    "plot raypath through system using array trace" 
    import matplotlib.pylab as plt
    import pyzdde.zdde as pyz
    ln = pyz.createLink()
    # launch rays from same from off-axis field point
    # we create initial pos and dir using zGetTraceArray
    nRays=7;  
    startsurf= 1;  # in case of collimated input beam    
    lastsurf = ln.zGetNumSurf();
    hx,hy,px,py = 0, 0.5, 0, _np.linspace(-1,1,nRays);
    (_,_,pos,dir,_,_) = zGetTraceArray(hx,hy,px,py,bParaxial=False,surf=startsurf);    
    # trace ray until last surface
    points = _np.zeros((lastsurf+1,nRays,3));    # indexing: surf,ray,coord
    z0=0; points[startsurf]=pos;                # ray intersection points on starting surface
    for isurf in range(startsurf,lastsurf):
      # trace to next surface
      (error,vigcode,pos,dir,_,_)=zGetTraceDirectArray(pos,dir,bParaxial=False,startSurf=isurf,lastSurf=isurf+1);
      points[isurf+1]=pos;
      points[isurf+1,vigcode!=0]=_np.nan;        # remove vignetted rays
      # add thickness of current surface (assumes absence of tilts or decenters in system)      
      z0+=ln.zGetThickness(isurf);
      points[isurf+1,:,2]+=z0;
      print("  surface #%d at z-position z=%f" % (isurf+1,z0));
    # plot rays in y-z section
    plt.figure();
    x,y,z = points[startsurf:].T;
    ax=plt.subplot(111,aspect='equal')
    ax.plot(z.T,y.T,'.-')
    ln.close();      


if __name__ == '__main__':
    # run the test functions
    _test_zGetTraceArray()
    _test_zGetTraceArrayDirect()
    _test_zOPDArray()
    _plot_2D_layout_from_array_trace()