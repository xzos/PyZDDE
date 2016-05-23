# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ArrayTraceUnittest.py
# Purpose:     unit tests for ArrayTrace
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

import pyzdde.zdde as pyzdde
import pyzdde.arraytrace as at
from   test.pyZDDEunittest import get_test_file

class TestArrayTrace(unittest.TestCase):

  @classmethod  
  def setUpClass(self):
    print('RUNNING TESTS FOR MODULE \'%s\'.'% at.__file__)
    self.ln = pyzdde.createLink();
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
    
  def test_zArrayTrace(self):
    print("\nTEST: arraytrace.zArrayTrace()")
    # Load a lens file into the Lde
    filename = get_test_file()
    self.ln.zLoadFile(filename)    
    self.ln.zPushLens(1);
    # set up field and pupil sampling
    nr = 9
    rd = at.getRayDataArray(nr)
    k = 0
    for i in xrange(-10, 11, 10):
        for j in xrange(-10, 11, 10):
            k += 1
            rd[k].z = i/20.0                   # px
            rd[k].l = j/20.0                   # py
            rd[k].intensity = 1.0
            rd[k].wave = 1
            rd[k].want_opd = 0
    # run array trace (C-extension)
    ret = at.zArrayTrace(rd)
    self.assertEqual(ret,0);
    #r = rd[1]; # select first ray   
    #for key in ('error','vigcode','x','y','z','l','m','n','Exr','Eyr','Ezr','opd','intensity'):
    #  print('self.assertAlmostEqual(rd[1].%s,%.8g,msg=\'%s differs\');'%(key,getattr(rd[1],key),key));
    self.assertAlmostEqual(rd[1].error,0,msg='error differs');
    self.assertAlmostEqual(rd[1].vigcode,0,msg='vigcode differs');
    self.assertAlmostEqual(rd[1].x,-0.0029856861,msg='x differs');
    self.assertAlmostEqual(rd[1].y,-0.0029856861,msg='y differs');
    self.assertAlmostEqual(rd[1].z,0,msg='z differs');
    self.assertAlmostEqual(rd[1].l,0.050136296,msg='l differs');
    self.assertAlmostEqual(rd[1].m,0.050136296,msg='m differs');
    self.assertAlmostEqual(rd[1].n,0.99748318,msg='n differs');
    self.assertAlmostEqual(rd[1].Exr,0,msg='Exr differs');
    self.assertAlmostEqual(rd[1].Eyr,0,msg='Eyr differs');
    self.assertAlmostEqual(rd[1].Ezr,-1,msg='Ezr differs');
    self.assertAlmostEqual(rd[1].opd,64.711234,msg='opd differs',places=5);
    self.assertAlmostEqual(rd[1].intensity,1,msg='intensity differs');


    
  def test_zGetTraceNumpy(self):
    print("\nTEST: arraytrace.zGetTraceNumpy()")
    # Load a lens file into the LDE
    filename = get_test_file()
    self.ln.zLoadFile(filename)
    self.ln.zPushLens(1);  
    # set-up field and pupil sampling
    x = np.linspace(-1,1,3)
    px= np.linspace(-1,1,3)    
    grid = np.meshgrid(x,x,px,px);
    field= np.transpose(grid[0:2]).reshape(-1,2);
    pupil= np.transpose(grid[2:4]).reshape(-1,2);
    # array trace (C-extension)
    ret = at.zGetTraceNumpy(field,pupil,mode=0);
    self.assertEqual(len(field),3**4);
        
    #for i in xrange(len(ret)):
    #  name = ['error','vigcode','pos','dir','normal','opd','intensity']
    #  print('self.assertAlmostEqual(ret[%d][1],%s,msg=\'%s differs\');'%(i,str(ret[i][1]),name[i]));
    self.assertEqual(ret[0][1],0,msg='error differs');
    self.assertEqual(ret[1][1],3,msg='vigcode differs');
    self.assertTrue(np.allclose(ret[2][1],[-18.24210131, -0.0671553, 0.]),msg='pos differs');
    self.assertTrue(np.allclose(ret[3][1],[-0.24287826, 0.09285061, 0.96560288]),msg='dir differs');
    self.assertTrue(np.allclose(ret[4][1],[ 0, 0, -1]),msg='normal differs');
    self.assertAlmostEqual(ret[5][1],66.8437599679,msg='opd differs');
    self.assertAlmostEqual(ret[6][1],1.0,msg='intensity differs');
    

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
      rd[k+1].want_opd = 0
    # results of zArrayTrace  
    ret = at.zArrayTrace(rd);
    self.assertEqual(ret,0);
    results = np.asarray( [[r.error,r.vigcode,r.x,r.y,r.z,r.l,r.m,r.n,\
                             r.Exr,r.Eyr,r.Ezr,r.opd,r.intensity] for r in rd[1:]] );
    # results of GetTraceArray
    (error,vigcode,pos,dir,normal,opd,intensity) = \
        at.zGetTraceNumpy(field,pupil,mode=0);

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
