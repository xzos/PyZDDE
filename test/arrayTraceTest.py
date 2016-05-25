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
import os
import sys
import unittest
import numpy as np

import pyzdde.zdde as pyz
import pyzdde.arraytrace as at
import pyzdde.arraytrace.numpy_interface as nt

from   test.pyZDDEunittest import get_test_file

class TestArrayTrace(unittest.TestCase):

  @classmethod  
  def setUpClass(self):
    print('RUNNING TESTS FOR MODULE \'%s\'.'% at.__file__)
    self.ln = pyz.createLink();
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
    
  def test_zGetTraceArray(self):
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
      for i in xrange(nr):
        reference = self.ln.zGetTrace(w,mode,-1,hx[i],hy[i],px[i],py[i]); 
        is_close = np.isclose(ret[:,i], np.asarray(reference));
        msg = 'zGetTraceArray differs from GetTrace for %s ray #%d:\n' % (mode_descr[mode],i);
        msg+= '  field: (%f,%f), pupil: (%f,%f) \n' % (hx[i],hy[i],px[i],py[i]);
        msg+= '  parameter  zGetTraceArray  zGetTrace \n';
        for j in np.arange(12)[~is_close]:
          msg+= '%10s  %12g  %12g \n'%(ret_descr[j],ret[j,i],reference[j]);
        self.assertTrue(np.all(is_close), msg=msg);


    
  def test_zGetTraceNumpy(self):
    print("\nTEST: arraytrace.numpy_interface.zGetTraceArray()")
    # Load a lens file into the LDE
    filename = get_test_file()
    self.ln.zLoadFile(filename)
    self.ln.zPushLens(1);  
    # set-up field and pupil sampling
    nr = 22; np.random.seed(0);    
    field = 2*np.random.rand(nr,2)-1; hx,hy = field.T;
    pupil = 2*np.random.rand(nr,2)-1; px,py = pupil.T;
    w = 1; # wavenum    
    
    # run array trace (C-extension), returns (error,vigcode,pos,dir,normal,opd,intensity)
    mode_descr = ("real","paraxial")    
    for mode in (0,1):
      print("  compare with GetTrace for %s raytrace"% mode_descr[mode]);
      ret = nt.zGetTraceArray(field,pupil,bParaxial=(mode==1),waveNum=w,surf=-1);
      mask = np.ones(13,dtype=np.bool); mask[-2]=False;   # mask array for removing opd from ret  
      ret = np.column_stack(ret).T;
      ret = ret[mask];
             
      # compare with results from GetTrace, returns (error,vig,x,y,z,l,m,n,l2,m2,n2,intensity)
      ret_descr = ('error','vigcode','x','y','z','l','m','n','Exr','Eyr','Ezr','intensity')
      for i in xrange(nr):
        reference = self.ln.zGetTrace(w,mode,-1,hx[i],hy[i],px[i],py[i]); 
        is_close = np.isclose(ret[:,i], np.asarray(reference));
        msg = 'zGetTraceArray differs from GetTrace for %s ray #%d:\n' % (mode_descr[mode],i);
        msg+= '  field: (%f,%f), pupil: (%f,%f) \n' % (hx[i],hy[i],px[i],py[i]);
        msg+= '  parameter  zGetTraceArray  zGetTrace \n';
        for j in np.arange(12)[~is_close]:
          msg+= '%10s  %12g  %12g \n'%(ret_descr[j],ret[j,i],reference[j]);
        self.assertTrue(np.all(is_close), msg=msg);


  def test_cross_check_zArrayTrace_vs_zGetTraceNumpy(self):
    print("\nTEST: comparison of zArrayTrace and zGetTraceNumpy")
    # Load a lens file into the LDE
    filename = get_test_file()
    self.ln.zLoadFile(filename)
    self.ln.zPushLens(1);  
    # set-up field and pupil sampling
    nr = 22;
    rd = at.getRayDataArray(nr)
    pupil = 2*np.random.rand(nr,2)-1;
    field = 2*np.random.rand(nr,2)-1;
    for k in xrange(nr):
      rd[k+1].x = field[k,0];
      rd[k+1].y = field[k,1];
      rd[k+1].z = pupil[k,0];
      rd[k+1].l = pupil[k,1];
      rd[k+1].intensity = 1.0;
      rd[k+1].wave = 1;
      rd[k+1].want_opd = -1;
    # results of zArrayTrace  
    ret = at.zArrayTrace(rd);
    self.assertEqual(ret,0);
    results = np.asarray( [[r.error,r.vigcode,r.x,r.y,r.z,r.l,r.m,r.n,\
                             r.Exr,r.Eyr,r.Ezr,r.opd,r.intensity] for r in rd[1:]] );
    # results of GetTraceArray
    (error,vigcode,pos,dir,normal,opd,intensity) = \
        nt.zGetTraceArray(field,pupil,bParaxial=0,want_opd=-1);

    # compare
    self.assertTrue(np.array_equal(error,results[:,0]),msg="error differs");    
    self.assertTrue(np.array_equal(vigcode,results[:,1]),msg="vigcode differs");    
    self.assertTrue(np.array_equal(pos,results[:,2:5]),msg="pos differs");    
    self.assertTrue(np.array_equal(dir,results[:,5:8]),msg="dir differs");    
    self.assertTrue(np.array_equal(normal,results[:,8:11]),msg="normal differs");    
    self.assertTrue(np.array_equal(opd,results[:,11]),msg="opd differs");    
    self.assertTrue(np.array_equal(intensity,results[:,12]),msg="intensity differs");    

    


if __name__ == '__main__':
  unittest.main()
