# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ArrayTraceUnittest.py
# Purpose:     unit tests for ArrayTrace
# Run:         from home directory of the project, i.e.,
#              $ python -m unittest test.arrayTraceTest
#
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
from __future__ import division
from __future__ import print_function

import unittest
import numpy as np

from _context import pyzdde
import pyzdde.zdde as pyz
import pyzdde.arraytrace as at
import pyzdde.arraytrace.numpy_interface as nt

from   .pyZDDEunittest import get_test_file

class TestArrayTrace(unittest.TestCase):

  @classmethod  
  def setUpClass(self):
    print('RUNNING TESTS FOR MODULE \'%s\'.'% at.__file__)
    self.ln = pyz.createLink();
    if self.ln is None:
      raise RuntimeError("Zemax DDE link could not be established. Please open Zemax.");
    self.ln.zGetUpdate();
   
  @classmethod
  def tearDownClass(self):
    self.ln.close();
    
    
  def test_getRayDataArray(self):
    print("\nTEST: arraytrace.getRayDataArray()")
    
    # create RayData without any kwargs
    rd = at.getRayDataArray(numRays=5)
    self.assertEqual(len(rd), 6)
    self.assertEqual(rd[0].error, 5)     # number of rays
    self.assertEqual(rd[0].opd, 0)       # GetTrace ray tracing type
    self.assertEqual(rd[0].wave, 0)      # real ray tracing
    self.assertEqual(rd[0].want_opd,-1)  # last surface

    # create RayData with some more arguments
    rd = at.getRayDataArray(numRays=5, tType=3, mode=1, startSurf=2)
    self.assertEqual(rd[0].opd, 3)       # mode 3
    self.assertEqual(rd[0].wave,1)       # real ray tracing
    self.assertEqual(rd[0].vigcode,2)    # first surface

    # create RayData with kwargs
    rd = at.getRayDataArray(numRays=5, tType=2, x=1.0, y=1.0)
    self.assertEqual(rd[0].x,1.0)
    self.assertEqual(rd[0].y,1.0)
    self.assertEqual(rd[0].z,0.0)

    # create RayData with kwargs overriding some regular parameters
    rd = at.getRayDataArray(numRays=5, tType=2, x=1.0, y=1.0, error=1)
    self.assertEqual(len(rd),6)
    self.assertEqual(rd[0].x,1.0)
    self.assertEqual(rd[0].y,1.0)
    self.assertEqual(rd[0].z,0.0)
    self.assertEqual(rd[0].error,1)
    
  def compare_array_with_single_trace(self,strace,atrace,param_descr,ret_descr,nr=22,seed=0):
    """ 
    helper function for comparing array raytrace and single raytrace functions
    for nr random rays which are constructed from given params (e.g. ['hx','hy','px','py']).
    The random number generator is initialized with given seed to ensure reproducibility.
    """
    # Load a lens file into the LDE
    filename = get_test_file()
    ret = self.ln.zLoadFile(filename);
    if ret!=0: raise IOError("Could not load Zemax file '%s'. Error code %d" % (filename,ret));
    if not self.ln.zPushLensPermission():
      raise RuntimeError("Extensions not allowed to push lenses. Please enable in Zemax.")
    self.ln.zPushLens(1);
    # set-up field and pupil sampling
    np.random.seed(seed);    
    params = 2*np.random.rand(len(param_descr),nr)-1;
    # perform array trace
    aret = atrace(*params)

    # compare with results from single raytrace
    for i in range(nr):
      sret = strace(*params[:,i]);
      is_close = np.isclose(aret[:,i], np.asarray(sret));
      msg = 'array and single raytrace differ for ray #%d:\n' % i;
      msg+= '  initial ray parameters: (%s)\n' % ",".join(param_descr);
      msg+= '                       ' + str(params[:,i]) + "\n";
      msg+= '  parameter  array-trace   single-trace \n';
      for j in np.arange(aret.shape[0])[~is_close]:
        msg+= '%10s  %12g  %12g \n'%(ret_descr[j],aret[j,i],sret[j]);
      self.assertTrue(np.all(is_close), msg=msg);   

  @unittest.skip("To be removed")
  def test_zGetTraceArrayOLd(self):
    print("\nTEST: arraytrace.zGetTraceArray()")
    # Load a lens file into the Lde
    filename = get_test_file()
    self.ln.zLoadFile(filename) 
    self.ln.zPushLens(1);
    # set up field and pupil sampling with random values
    nr = 22; np.random.seed(0);    
    hx,hy,px,py = 2*np.random.rand(4,nr)-1;
    w = 1; # wavenum    
    
    # run array trace (C-extension), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,opd,intensity)
    mode_descr = ("real","paraxial")    
    for mode in (0,1):
      print("  compare with GetTrace for %s raytrace"% mode_descr[mode]);
      ret = at.zGetTraceArray(nr,list(hx),list(hy),list(px),list(py),
                      intensity=1,waveNum=w,mode=mode,surf=-1,want_opd=0)
      mask = np.ones(13,dtype=np.bool); mask[-2]=False;   # mask array for removing opd from ret  
      ret = np.asarray(ret)[mask];
       
      # compare with results from GetTrace, returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','Exr','Eyr','Ezr','intensity')
      for i in range(nr):
        reference = self.ln.zGetTrace(w,mode,-1,hx[i],hy[i],px[i],py[i]); 
        is_close = np.isclose(ret[:,i], np.asarray(reference));
        msg = 'zGetTraceArray differs from GetTrace for %s ray #%d:\n' % (mode_descr[mode],i);
        msg+= '  field: (%f,%f), pupil: (%f,%f) \n' % (hx[i],hy[i],px[i],py[i]);
        msg+= '  parameter  zGetTraceArray  zGetTrace \n';
        for j in np.arange(12)[~is_close]:
          msg+= '%10s  %12g  %12g \n'%(ret_descr[j],ret[j,i],reference[j]);
        self.assertTrue(np.all(is_close), msg=msg);

  def test_cross_check_zArrayTrace_vs_zGetTraceNumpy(self):
    print("\nTEST: comparison zGetTraceArray from numpy_interface and raystruct_interface.")
    w = 1; # wavenum
    nr= 3; # number of rays
    
    for mode,descr in [(0,"real"),(1,"paraxial")]:
      print("  compare zGetTraceNumpy (called with single ray) with zGetTraceArray() for %s raytrace"% descr);
      # single trace (GetTrace), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','l2','m2','n2','intensity')
      def strace(hx,hy,px,py):
        ret = nt.zGetTraceArray(hx,hy,px,py,intensity=1,waveNum=w,bParaxial=(mode==1),surf=-1) 
        return np.column_stack(ret).flatten(); 
      # array trace (C-extension), returns (error,vigcode,pos(3),dir(3),normal(3)intensity)
      def atrace(hx,hy,px,py):
        ret = at.zGetTraceArray(nr,list(hx),list(hy),list(px),list(py),
                                intensity=1,waveNum=w,mode=mode,surf=-1,want_opd=0);
        mask = np.ones(13,dtype=np.bool); mask[-2]=False;   # mask array for removing opd from ret    
        return np.column_stack(ret).T[mask];
      # perform comparison  
      self.compare_array_with_single_trace(strace,atrace,('hx','hy','px','py'),ret_descr,nr=nr);
  
  @unittest.skip("To be removed")
  def test_cross_check_zArrayTrace_vs_zGetTraceNumpy_OLD(self):
    print("\nTEST: comparison of zArrayTrace and zGetTraceNumpy OLD")
    # Load a lens file into the LDE
    filename = get_test_file()
    self.ln.zLoadFile(filename)
    self.ln.zPushLens(1);  
    # set-up field and pupil sampling
    nr = 22;
    rd = at.getRayDataArray(nr)
    hx,hy,px,py = 2*np.random.rand(4,nr)-1;
    
    for k in range(nr):
      rd[k+1].x = hx[k];
      rd[k+1].y = hy[k];
      rd[k+1].z = px[k];
      rd[k+1].l = py[k];
      rd[k+1].intensity = 1.0;
      rd[k+1].wave = 1;
      rd[k+1].want_opd = 0;
    # results of zArrayTrace  
    ret = at.zArrayTrace(rd);
    self.assertEqual(ret,0);
    results = np.asarray( [[r.error,r.vigcode,r.x,r.y,r.z,r.l,r.m,r.n,\
                             r.Exr,r.Eyr,r.Ezr,r.opd,r.intensity] for r in rd[1:]] );
    # results of GetTraceArray
    (error,vigcode,pos,dir,normal,intensity) = \
        nt.zGetTraceArray(hx,hy,px,py,bParaxial=0);

    # compare
    self.assertTrue(np.array_equal(error,results[:,0]),msg="error differs");    
    self.assertTrue(np.array_equal(vigcode,results[:,1]),msg="vigcode differs");    
    self.assertTrue(np.array_equal(pos,results[:,2:5]),msg="pos differs");    
    self.assertTrue(np.array_equal(dir,results[:,5:8]),msg="dir differs");    
    self.assertTrue(np.array_equal(normal,results[:,8:11]),msg="normal differs");    
    self.assertTrue(np.array_equal(intensity,results[:,12]),msg="intensity differs");       

  def test_zGetTraceNumpy(self):
    print("\nTEST: arraytrace.numpy_interface.zGetTraceArray()")
    w = 1; # wavenum
    
    for mode,descr in [(0,"real"),(1,"paraxial")]:
      print("  compare with GetTrace for %s raytrace"% descr);
      # single trace (GetTrace), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','l2','m2','n2','intensity')
      def strace(hx,hy,px,py):
        return self.ln.zGetTrace(w,mode,-1,hx,hy,px,py); 
      # array trace (C-extension), returns (error,vigcode,pos(3),dir(3),normal(3)intensity)
      def atrace(hx,hy,px,py):
        ret = nt.zGetTraceArray(hx,hy,px,py,bParaxial=(mode==1),waveNum=w,surf=-1); 
        return np.column_stack(ret).T;
      # perform comparison  
      self.compare_array_with_single_trace(strace,atrace,('hx','hy','px','py'),ret_descr);
      
      
  def test_zGetOpticalPathDifference(self):
    print("\nTEST: arraytrace.numpy_interface.zGetOpticalPathDifference()")
    w = 1; # wavenum
    px,py=1,0.5;  # we fix the pupil values, as GetOpticalPathDifference
                  # traces rays to all pupil points for each field point
    # single trace (GetTrace,OPDX), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
    ret_descr = ('error','vigcode','opd','x','y','z','l','m','n','intensity')
    def strace(hx,hy):
      (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)=self.ln.zGetTrace(w,0,-1,hx,hy,px,py);  # real ray trace to image surface
      opd=self.ln.zGetOpticalPathDifference(hx,hy,px,py,ref=0,wave=w);                   # calculate OPD, ref: chief ray
      vig=1 if vig!=0 else 0;    # vignetting flag is only 0 or 1 in ArrayTrace, not the surface number
      return (error,vig,opd,x,y,z,l,m,n,intensity);
    # array trace (C-extension), returns (error,vigcode,opd,pos,dir,intensity) 
    def atrace(hx,hy):
      ret = nt.zGetOpticalPathDifferenceArray(hx,hy,px,py,waveNum=w);
      ret = [ var.reshape((ret[0].size,-1)) for var in ret ]; # reshape each argument as (nRays,...)
      return np.hstack(ret).T;
    # perform comparison  
    self.compare_array_with_single_trace(strace,atrace,('hx','hy'),ret_descr);

  # -----------------------------------------------------------------------------
  # Test works with Zemax 13 R2
  # Test fails with OpticStudio (ZOS16.5). We obtain different error and vigcodes 
  # using either single raytrace or arraytrace. Should be handled in the python 
  # interface to avoid confusion
  # -----------------------------------------------------------------------------
  def test_zGetTraceDirectNumpy(self):
    print("\nTEST: arraytrace.numpy_interface.zGetTraceDirectArray()")
    w = 1; # wavenum
    startSurf=0
    lastSurf=-1
    
    for mode,descr in [(0,"real"),(1,"paraxial")]:
      print("  compare with GetTraceDirect for %s raytrace"% descr);
      # single trace (GetTraceDirect), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','l2','m2','n2','intensity')
      def strace(x,y,z,l,m):
        n = np.sqrt(1-0.5*l**2-0.5*m**2); # calculate z-direction (scale l,m to < 1/sqrt(2))
        return self.ln.zGetTraceDirect(w,mode,startSurf,lastSurf,x,y,z,l,m,n); 
      # array trace (C-extension), returns (error,vigcode,pos(3),dir(3),normal(3)intensity)
      def atrace(x,y,z,l,m):
        n = np.sqrt(1-0.5*l**2-0.5*m**2); # calculate z-direction        
        pos = np.stack((x,y,z),axis=1);
        dir = np.stack((l,m,n),axis=1);
        ret = nt.zGetTraceDirectArray(pos,dir,bParaxial=mode,startSurf=startSurf,lastSurf=lastSurf,
                                      intensity=1,waveNum=w);
        return np.column_stack(ret).T;
      # perform comparison  
      self.compare_array_with_single_trace(strace,atrace,('x','y','z','l','m'),ret_descr);

  # -----------------------------------------------------------------------------
  # Test works with Zemax 13 R2
  # Test fails with OpticStudio (ZOS16.5). We obtain different error and vigcodes 
  # using either single raytrace or arraytrace. Should be handled in the python 
  # interface to avoid confusion
  # -----------------------------------------------------------------------------
  def test_zGetTraceDirectRaystruct(self):
    print("\nTEST: arraytrace.raystruct_interface.zGetTraceDirectArray()")
    w = 1; # wavenum
    startSurf=0
    lastSurf=-1
    
    for mode,descr in [(0,"real"),(1,"paraxial")]:
      print("  compare with GetTraceDirect for %s raytrace"% descr);
      # single trace (GetTraceDirect), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','l2','m2','n2','intensity')
      def strace(x,y,z,l,m):
        n = np.sqrt(1-0.5*l**2-0.5*m**2); # calculate z-direction (scale l,m to < 1/sqrt(2))
        return self.ln.zGetTraceDirect(w,mode,startSurf,lastSurf,x,y,z,l,m,n); 
      # array trace (C-extension), returns (error,vigcode,x,y,z,l,m,n,l2,m2,n2,opd,intensity)
      def atrace(x,y,z,l,m):
        n = np.sqrt(1-0.5*l**2-0.5*m**2); # calculate z-direction  
        ret = at.zGetTraceDirectArray(x.size,list(x),list(y),list(z),list(l),list(m),list(n),
                                      mode=mode,startSurf=startSurf,lastSurf=lastSurf, 
                                      intensity=1,waveNum=w);
        mask = np.ones(13,dtype=np.bool); mask[-2]=False;   # mask array for removing opd from ret    
        return np.column_stack(ret).T[mask];
      # perform comparison  
      self.compare_array_with_single_trace(strace,atrace,('x','y','z','l','m'),ret_descr);
  # -----------------------------------------------------------------------------
  # test fails for real raytrace, as arrayTrace from Zemax does not return
  # surface normal correctly. Should be handled in the python interface
  # to avoid confusion
  # -----------------------------------------------------------------------------
  def test_zGetTraceArrayOPD(self):
    print("\nTEST: OPD for arraytrace.raystruct_interface.zGetTraceArray()")
    w = 1; # wavenum
    nr=30; # number of rays
    mode=0;# only real raytrace works with want_opd
    
    # single trace (GetTrace), returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
    ret_descr = ('error','vigcode','x','y','z','l','m','n','l2','m2','n2','opd','intensity')
    def strace(hx,hy,px,py):
      opd=self.ln.zGetOpticalPathDifference(hx,hy,px,py,ref=0,wave=w);     
      (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity) = self.ln.zGetTrace(w,mode,-1,hx,hy,px,py); 
      return (error,vig,x,y,z,l,m,n,l2,m2,n2,opd,intensity);
    # array trace (C-extension), returns (error,vigcode,pos(3),dir(3),normal(3)intensity)
    def atrace(hx,hy,px,py):
      ret = at.zGetTraceArray(nr,list(hx),list(hy),list(px),list(py),
                              intensity=1,waveNum=w,mode=mode,surf=-1,want_opd=-1);
      return np.column_stack(ret).T;
    # perform comparison  
    self.compare_array_with_single_trace(strace,atrace,('hx','hy','px','py'),ret_descr,nr=nr);

    


if __name__ == '__main__':
  # see https://docs.python.org/2/library/unittest.html
  unittest.main(module='arrayTraceTest');
                             
