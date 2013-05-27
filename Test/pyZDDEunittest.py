#-------------------------------------------------------------------------------
# Name:        pyZDDEunittest.py
# Purpose:     pyZDDE unit test using the python unittest framework
#
# Author:      Indranil Sinharoy
#
# Created:     19/10/2012
# Copyright:   (c) Indranil Sinharoy, 2012 - 2013
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.2
#-------------------------------------------------------------------------------
from __future__ import division
from __future__ import print_function
import os
import sys
import unittest

# Put both the "Test" and the "PyZDDE" directory in the python search path.
testdirectory = os.path.dirname(os.path.realpath(__file__))
ind = testdirectory.find('Test')
pyzddedirectory = testdirectory[0:ind-1]
if testdirectory not in sys.path:
    sys.path.append(testdirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

import pyZDDE

reload(pyZDDE)  # In order to ensure that the latest changes in the pyZDDE module
                # are updated here.

# ZEMAX file directory
zmxfp = pyzddedirectory+'\\ZMXFILES\\'

# Zemax file(s) used in the test
lensFile = ["Cooke 40 degree field.zmx"]
lensFileName = lensFile[0]

# Flag to enable printing of returned values.
PRINT_RETURNED_VALUES = 1     # if test results are not going to be viewed by
                              # humans, turn this off.

class TestPyZDDEFunctions(unittest.TestCase):
    pRetVar = PRINT_RETURNED_VALUES

    def setUp(self):
        # Create the DDE channel object
        self.link0 = pyZDDE.pyzdde()
        # Initialize the DDE
        # The DDE initialization has be done here, and so cannot be tested
        # otherwise as no zDDExxx functions can be carried before initialization.
        status = self.link0.zDDEInit()
        if TestPyZDDEFunctions.pRetVar:
            print("Status for link 0:", status)
        self.assertEqual(status,0)
        if status:
            print("Couldn't initialize DDE.")
        #Make sure to reset the lens
        ret = self.link0.zNewLens()

    def tearDown(self):
        # Tear down unit test
        if self.link0.connection:
            self.link0.zDDEClose()
        else:
            print("Server was already terminated")

    @unittest.skip("Init is now being called in the setup")
    def test_zDDEInit(self):
        # Test initialization
        print("\nTEST: zDDEInit()")
        status = self.link0.zDDEInit()
        print("Status for link 0:", status)
        self.assertEqual(status,0)
        if status ==0:
            TestPyZDDEFunctions.b_zDDEInit_test_done = True
        else:
            print("Couln't initialize DDE.")

    @unittest.skip("Being called in the teardown")
    def test_zDDEClose(self):
        print("\nTEST: zDDEClose()")
        stat = self.link0.zDDEClose()

    @unittest.skip("To implement test")
    def test_zCloseUDOData(self):
        print("\nTEST: zCloseUDOData()")
        pass

    @unittest.skip("To implement test")
    def test_zDeleteMFO(self):
        print("\nTEST: zDeleteMFO()")
        pass

    @unittest.skip("To implement test")
    def test_zDeleteObject(self):
        print("\nTEST: zDeleteObject()")
        pass

    def test_zDeleteConfig(self):
        print("\nTEST: zDeleteConfig()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        currConfig = self.link0.zGetConfig()
        #Since no configuration is initally present, it should return (1,1,1)
        self.assertTupleEqual(currConfig,(1,1,1))
        #Insert a config
        self.link0.zInsertConfig(currConfig[1]+1)
        #Assert if the number of configurations didn't increase, however the
        #current configuration shouldn't change, and the number of multiple
        #configurations must remain same.
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1]+1,currConfig[2]))
        #Now, finally, call zDeleteConfig() to switch configuration
        configNum = self.link0.zDeleteConfig(2)
        self.assertEqual(configNum,2)
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,currConfig)

    def test_zDeleteMCO(self):
        print("\nTEST: zDeleteMCO()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Get the current number of configurations (columns and rows)
        currConfig = self.link0.zGetConfig()
        self.assertTupleEqual(currConfig,(1,1,1))
        #Insert a operand (row)
        newOperNumber = self.link0.zInsertMCO(2)
        self.assertEqual(newOperNumber,2)
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1],currConfig[2]+1))
        #Finally delete an MCO
        newOperNumber = self.link0.zDeleteMCO(2)
        self.assertEqual(newOperNumber,1)

    @unittest.skip("To implement test")
    def test_zDeleteSurface(self):
        print("\nTEST: zDeleteSurface()")
        pass

    @unittest.skip("To implement test")
    def test_zExportCAD(self):
        print("\nTEST: zExportCAD()")
        pass

    @unittest.skip("To implement test")
    def test_zExportCheck(self):
        print("\nTEST: zExportCheck()")
        pass

    @unittest.skip("To implement")
    def test_zFindLabel(self):
        print("\nTEST: zFindLabel()")
        pass

    @unittest.skip("To implement")
    def test_zGetAddress(self):
        print("\nTEST: zGetAddress()")
        pass

    @unittest.skip("To implement")
    def test_zGetAperture(self):
        print("\nTEST: zGetAperture()")
        pass

    @unittest.skip("To implement")
    def test_zGetApodization(self):
        print("\nTEST: zGetApodization()")
        pass

    @unittest.skip("To implement")
    def test_zGetAspect(self):
        print("\nTEST: zGetAspect()")
        pass

    @unittest.skip("To implement")
    def test_zGetBuffer(self):
        print("\nTEST: zGetBuffer()")
        pass

    @unittest.skip("Not important")
    def test_zGetComment(self):
        print("\nTEST: zGetComment()")
        pass

    def test_zGetConfig(self):
        print("\nTEST: zGetConfig()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        currConfig = self.link0.zGetConfig()
        #Since no configuration is initally present, it should return (1,1,1)
        self.assertTupleEqual(currConfig,(1,1,1))
        #Insert a config
        self.link0.zInsertConfig(currConfig[1]+1)
        #Assert if the number of configurations didn't increase, however the
        #current configuration shouldn't change, and the number of multiple
        #configurations must remain same.
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1]+1,currConfig[2]))
        if TestPyZDDEFunctions.pRetVar:
            print("CONFIG: ", newConfig)

    def test_zGetDate(self):
        print("\nTEST: zGetDate()")
        date = self.link0.zGetDate().rstrip()
        self.assertIsInstance(date,str)
        if TestPyZDDEFunctions.pRetVar:
            print("DATE: ", date)

    @unittest.skip("To implement")
    def test_zGetExtra(self):
        print("\nTEST: zGetExtra()")


    def test_zGetField(self):
        print("\nTEST: zGetField()")
        # First set some valid field parameters in the ZEMAX DDE server
        # set field with 4 arguments, n=0, 3 field points
        fieldData = self.link0.zSetField(0,0,3,1)
        # Set the first, second and third field point
        fieldData = self.link0.zSetField(1,0,0)
        fieldData = self.link0.zSetField(2,0,5,2.0,0.5,0.5,0.5,0.5,0.5)
        fieldData = self.link0.zSetField(3,0,10,1.0,0.0,0.0,0.0)
        # Now, verify the set parameters using zGetField()
        fieldData = self.link0.zGetField(0)
        self.assertTupleEqual((0,3),(fieldData[0],fieldData[1]))
        fieldData = self.link0.zGetField(1)
        self.assertTupleEqual(fieldData,(0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0))
        fieldData = self.link0.zGetField(2)
        self.assertTupleEqual(fieldData,(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5))
        fieldData = self.link0.zGetField(3)
        self.assertTupleEqual(fieldData,(0.0, 10.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0))

    def test_zGetFieldTuple(self):
        print("\nTEST: zGetFieldTuple()")
        # First set the field using setField tuple data in the ZEMAX server
        iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                              # vdx=vdy=vcx=vcy=van=0.0
                           (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                           (0.0,10.0))                        # field3: xf=0.0,yf=10.0
        # Set the field data, such that fieldType is angle with rectangular normalization
        oFieldDataTuple_s = self.link0.zSetFieldTuple(0,1,iFieldDataTuple)
        # Now get the field data by callling zGetFieldTuple
        oFieldDataTuple_g = self.link0.zGetFieldTuple()
        if TestPyZDDEFunctions.pRetVar:
            for i in range(len(iFieldDataTuple)):
                print("oFieldDataTuple_g, field {i} : {t}".format(i=i,
                                                        t=oFieldDataTuple_g[i]))
        #Verify
        for i in range(len(iFieldDataTuple)):
            self.assertEqual(oFieldDataTuple_g[i][:len(iFieldDataTuple[i])],
                                                       iFieldDataTuple[i])

    def test_zGetFile(self):
        print("\nTEST: zGetFile()")
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0   # This is not a unit-test assert.
        reply = self.link0.zGetFile()
        self.assertEqual(reply,filename)
        if TestPyZDDEFunctions.pRetVar:
            print("zGetFile return value: {}".format(reply))

    def test_zGetFirst(self):
        print("\nTEST: zGetFirst()")
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0   # This is not a unit-test assert.
        (focal,pwfn,rwfn,pima,pmag) = self.link0.zGetFirst()
        #Just going to check the validity of the returned data type
        self.assertIsInstance(focal,float)
        self.assertIsInstance(pwfn,float)
        self.assertIsInstance(rwfn,float)
        self.assertIsInstance(pima,float)
        if TestPyZDDEFunctions.pRetVar:
            print(("zGetFirst ret: {:.4f},{:.4f},{:.4f},{:.4f},{:.4f}"
                   .format(focal,pwfn,rwfn,pima,pmag)))

    @unittest.skip("To implement")
    def test_zGetGlass(self):
        print("\nTEST: zGetGlass()")
        pass

    @unittest.skip("To implement")
    def test_zGlobalMatrix(self):
        print("\nTEST: zGetGlobalMatrix()")
        pass

    @unittest.skip("To implement")
    def test_zGetIndex(self):
        print("\nTEST: zGetIndex()")
        pass

    @unittest.skip("To implement")
    def test_zGetLabel(self):
        print("\nTEST: zGetLabel()")
        pass

    @unittest.skip("To implement")
    def test_zGetMetaFile(self):
        print("\nTEST: zGetMetaFile()")
        pass

    @unittest.skip("To implement test")
    def test_zGetMode(self):
        print("\nTEST: zGetMode()")
        pass

    def test_zGetMulticon(self):
        print("\nTEST: zGetMulticon()")
        #Test zGetMulticon return when the MCE is "empty" (it shouldn't error out)
        multiConData = self.link0.zGetMulticon(2,3)  # configuration 2, row 3 (both are fictitious)
        self.assertIsInstance(multiConData,tuple)
        #insert an additional configuration (column)
        self.link0.zInsertConfig(1)
        #insert an additional operand (row)
        self.link0.zInsertMCO(2)
        #Set the row operands (both to thickness, of surfaces 2, and 4 respectively)
        multiConData = self.link0.zSetMulticon(0,1,'THIC',2,0,0)
        multiConData = self.link0.zSetMulticon(0,2,'THIC',4,0,0)
        #Set configuration 1
        multiConData = self.link0.zSetMulticon(1,1,6.0076,0,1,1,1.0,0.0)
        multiConData = self.link0.zSetMulticon(1,2,4.7504,0,1,1,1.0,0.0)
        #Set configuration 2
        multiConData = self.link0.zSetMulticon(2,1,7.0000,0,1,1,1.0,0.0)
        multiConData = self.link0.zSetMulticon(2,2,5.0000,0,1,1,1.0,0.0)
        #use zGetMulticon() to verify the set values
        multiConData = self.link0.zGetMulticon(1,1) # row 1, config 1
        self.assertTupleEqual(multiConData,(6.0076, 2, 2, 0, 1, 1, 1.0, 0.0))
        multiConData = self.link0.zGetMulticon(2,1) # row 1, config 2
        self.assertTupleEqual(multiConData,(7.0, 2, 2, 0, 1, 1, 1.0, 0.0))
        multiConData = self.link0.zGetMulticon(1,2) # row 2, config 1
        self.assertTupleEqual(multiConData,(4.7504, 2, 2, 0, 1, 1, 1.0, 0.0))
        multiConData = self.link0.zGetMulticon(2,2) # row 2, config 2
        self.assertTupleEqual(multiConData,(5.0, 2, 2, 0, 1, 1, 1.0, 0.0))

    @unittest.skip("To implement test")
    def test_zGetName(self):
        print("\nTEST: zGetName()")
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        reply = self.lin0.zGetName()
        self.assertEqual(reply,"A SIMPLE COOKE TRIPLET.")

    @unittest.skip("To implement test")
    def test_zGetNSCData(self):
        print("\nTEST: zGetNSCData()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCMatrix(self):
        print("\nTEST: zGetNSCMatrix()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCObjectData(self):
        print("\nTEST: zGetNSCObjectData()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCObjectFaceData(self):
        print("\nTEST: zGetNSCObjectFaceData()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCParameter(self):
        print("\nTEST: zGetNSCParameter()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCPosition(self):
        print("\nTEST: zGetNSCPosition()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCProperty(self):
        print("\nTEST: zGetNSCProperty()")
        pass

    @unittest.skip("To implement test")
    def test_zGetNSCSetting(self):
        print("\nTEST: zGetNSCSetting()")
        pass

    @unittest.skip("To implement test")
    def test_zGetOperand(self):
        print("\nTEST: zGetOperand()")
        pass

    def test_zGetPath(self):
        print("\nTEST: zGetPath()")
        (p2DataFol,p2DefaultFol) = self.link0.zGetPath()
        self.assertTrue(os.path.isabs(p2DataFol))
        self.assertTrue(os.path.isabs(p2DefaultFol))

    def test_zGetPolState(self):
        print("\nTEST: zGetPolState()")
        #Set polarization of the "new" lens
        self.link0.zSetPolState(0,0.5,0.5,10.0,10.0)
        polStateData = self.link0.zGetPolState()
        self.assertTupleEqual(polStateData,(0,0.5,0.5,10.0,10.0))

    def test_zGetPolTrace(self):
        print("\nTEST: zGetPolTrace()")
       # Load a lens file into the LDE
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Set up the data
        (waveNum,mode,surf,hx,hy,px,py,Ex,Ey,Phax,Phay) = (1,0,-1,0.0,.5,0.0,1.0,
                                                        0.7071067,0.7071067,0,0)
        rayTraceArg = (waveNum,mode,surf,hx,hy,px,py,Ex,Ey,Phax,Phay)
        expRayTraceData = (0, 0.9403035211373325, -0.3816204067506139, 0.0, -0.0,
                           0.89144230676406, 0.0, 0.0)
        # test returned tuple
        rayTraceData = self.link0.zGetPolTrace(*rayTraceArg)
        for i,d in enumerate(expRayTraceData):
            self.assertAlmostEqual(rayTraceData[i],d,places=4)

    def test_zGetPupil(self):
        print("\nTEST: zGetPupil()")
        # Load a lens to the ZEMAX DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Get the pupil data
        pupilData = self.link0.zGetPupil()
        expPupilData = (0, 10.0, 10.0, 11.51215705, 10.23372788, -50.96133853, 0, 0.0)
        for i,d in enumerate(expPupilData):
            self.assertAlmostEqual(pupilData[i],d,places=4)
        #Print the pupil data if switch is on.
        if TestPyZDDEFunctions.pRetVar:
            pupil_data = dict(zip((0,1,2,3,4,5,6,7),('type','value','ENPD','ENPP',
                       'EXPD','EXPP','apodization_type','apodization_factor')))
            pupil_type = dict(zip((0,1,2,3,4,5),
                ('entrance pupil diameter','image space F/#','object space NA',
                  'float by stop','paraxial working F/#','object cone angle')))
            pupil_value_type = dict(zip((0,1),("stop surface semi-diameter",
                                             "system aperture")))
            apodization_type = dict(zip((0,1,2),('none','Gaussian','Tangential')))
            # Print the pupil data
            print("Pupil data:")
            print("{pT} : {pD}".format(pT=pupil_data[0],pD=pupil_type[pupilData[0]]))
            print("{pT} : {pD} {pV}".format(pT = pupil_data[1], pD=pupilData[1],
                                            pV = (pupil_value_type[0]
                                            if pupilData[0]==3 else
                                            pupil_value_type[1])))
            for i in range(2,6):
                print("{pd} : {pD:2.4f}".format(pd=pupil_data[i],pD=pupilData[i]))
            print("{pd} : {pD}".format(pd=pupil_data[6],pD=apodization_type[pupilData[6]]))
            print("{pd} : {pD:2.4f}".format(pd=pupil_data[7],pD=pupilData[7]))

    def test_zGetRefresh(self):
        print("\nTEST: zGetRefresh()")
        # Load & then push a lens file into the LDE
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        ret = self.link0.zPushLens(1)
        # Copy the lens data from the LDE into the stored copy of the ZEMAX server.
        ret = self.link0.zGetRefresh()
        self.assertIn(ret,(-998,-1,0))
        if ret == -1:
            print("MSG: ZEMAX couldn't copy the lens data to the LDE")
        if ret == -998:
            print("MSG: zGetRefresh() function timed out")
        if TestPyZDDEFunctions.pRetVar:
            print("zGetRefresh return value", ret)

    @unittest.skip("To implement")
    def test_zGetSag(self):
        print("\nTEST: zGetSag()")
        #Load a lens file

    @unittest.skip("To implement")
    def test_zGetSequence(self):
        print("\nTEST: zGetSequence()")
        #Load a lens file

    def test_zGetSerial(self):
        print("\nTEST: zGetSerial()")
        ser = self.link0.zGetSerial()
        self.assertIsInstance(ser,int)
        if TestPyZDDEFunctions.pRetVar:
            print("SERIAL:", ser)

    @unittest.skip("To implement")
    def test_zGetSettingsData(self):
        print("\nTEST: zGetSettingsData()")

    @unittest.skip("To implement")
    def test_zGetSolve(self):
        print("\nTEST: zGetSolve()")

    def test_zGetSurfaceData(self):
        print("\nTEST: zGetSurfaceData()")
        #Load a lens file
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0
        surfName = self.link0.zGetSurfaceData(1,0)   # Surface name
        self.assertEqual(surfName,'STANDARD')
        radius = 1.0/self.link0.zGetSurfaceData(1,2) # curvature
        self.assertAlmostEqual(radius,22.01359,places=3)
        thick = self.link0.zGetSurfaceData(1,3)     # thickness
        self.assertAlmostEqual(thick,3.25895583,places=3)
        if TestPyZDDEFunctions.pRetVar:
            print("surfName :", surfName)
            print("radius :", radius)
            print("thickness: ", thick)
        #ToDo: call zGetSurfaceData() with 3 arguments

    @unittest.skip("To implement")
    def test_zGetSurfaceDLL(self):
        print("\nTEST: zGetSurfaceDLL()")
        #Load a lens file

    @unittest.skip("To implement")
    def test_zGetSurfaceParameter(self):
        print("\nTEST: zGetSurfaceParameter()")
        #Load a lens file
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0
        surfParam1 = self.link0.zGetSurfaceParameter(1,1)
        print("Surface name: ", surfParam1)
        surfParam3 = self.link0.zGetSurfaceParameter(1,3)
        print("Radius: ", surfParam3)
        #ToDo: not complete


    def test_zGetSystem(self):
        print("\nTEST: zGetSystem()")
        #Setup the arguments to set a specific system first
        unitCode,stopSurf,rayAimingType = 0,4,0  # mm, 4th,off
        useEnvData,temp,pressure,globalRefSurf = 0,20,1,1 # off, 20C,1ATM,ref=1st surf
        setSystemArg = (unitCode,stopSurf,rayAimingType,useEnvData,
                                                  temp,pressure,globalRefSurf)
        expSystemData = (2, 0, 2, 0, 0, 0, 20.0, 1, 1, 0)
        recSystemData_s = self.link0.zSetSystem(*setSystemArg)
        # Now get the system data using zGetSystem(), the returned structure
        # should be same as that returned by zSetSystem()
        recSystemData_g = self.link0.zGetSystem()
        self.assertTupleEqual(recSystemData_s,recSystemData_g)
        if TestPyZDDEFunctions.pRetVar:
            systemDataPar = ('numberOfSurfaces','lens unit code',
                             'stop surface-number','non axial flag',
                             'ray aiming type','adjust index','current temperature',
                             'pressure','global surface reference')  #'need_save' Deprecated.
            print("System data:")
            for i,elem in enumerate(systemDataPar):
                print("{el} : {sd}".format(el=elem,sd=recSystemData_g[i]))

    def test_zGetSystemApr(self):
        print("\nTEST: zGetSystemApr()")
        # First set the system aperture to known parameters in the ZEMAX server
        systemAperData_s = self.link0.zSetSystemAper(0,1,25) #sysAper=25mm,EPD
        systemAperData_g = self.link0.zGetSystemAper()
        self.assertTupleEqual(systemAperData_s,systemAperData_g)

    def test_zGetSystemProperty(self):
        print("\nTEST: zGetSystemProperty():")
        #Set Aperture type as EPD
        sysPropData_s = self.link0.zSetSystemProperty(10,0)
        sysPropData_g = self.link0.zGetSystemProperty(10)
        self.assertEqual(sysPropData_s,sysPropData_g)
        #Let lens title
        sysPropData_s = self.link0.zSetSystemProperty(16,"My Lens")
        sysPropData_g = self.link0.zGetSystemProperty(16)
        self.assertEqual(sysPropData_s,sysPropData_g)
        #Set glass catalog
        sysPropData_s = self.link0.zSetSystemProperty(23,"SCHOTT HOYA OHARA")
        sysPropData_g = self.link0.zGetSystemProperty(23)
        self.assertEqual(sysPropData_s,sysPropData_g)

    def test_zGetTextFile(self):
        print("\nTEST: zGetTextFile()")
        # Load a lens file into the DDE Server (Not required to Push lens)
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # create text files
        spotDiagFileName = 'SpotDiagram.txt'          # Change appropriately
        abberSCFileName = 'SeidelCoefficients.txt'    # Change appropriately
        #Request to dump prescription file, without giving fullpath name. It
        #should return -1
        preFileName = 'Prescription_unitTest_00.txt'
        ret = self.link0.zGetTextFile(preFileName,'Pre',"None",0)
        self.assertEqual(ret,-1)
        #filename path is absolute, however, doesn't have extension
        textFileName = testdirectory + '\\' + os.path.splitext(preFileName)[0]
        ret = self.link0.zGetTextFile(textFileName,'Pre',"None",0)
        self.assertEqual(ret,-1)
        #Request to dump prescription file, without providing a valid settings file
        #and flag = 0 ... so that the default settings will be used for the text
        #Create filename with full path
        textFileName = testdirectory + '\\' + preFileName
        ret = self.link0.zGetTextFile(textFileName,'Pre',"None",0)
        self.assertIn(ret,(0,-1,-998)) #ensure that the ret is any valid return
        if ret == -1:
            print("MSG: zGetTextFile failed")
        if ret == -998:
            print("MSG: zGetTextFile() function timed out")
        if TestPyZDDEFunctions.pRetVar:
            print("zGetTextFile return value", ret)
        #Request zemax to dump prescription file, with a settings
        ret = self.link0.zGetRefresh()
        settingsFileName = "Cooke 40 degree field_PreSettings_OnlyCardinals.CFG"
        preFileName = 'Prescription_unitTest_01.txt'
        textFileName = testdirectory + '\\' + preFileName
        ret = self.link0.zGetTextFile(textFileName,'Pre',settingsFileName,1)
        self.assertIn(ret,(0,-1,-998)) #ensure that the ret is any valid return
        if ret == -1:
            print("MSG: zGetText failed")
        if ret == -998:
            print("MSG: zGetText() function timed out")
        if TestPyZDDEFunctions.pRetVar:
            print("zGetText return value", ret)
        #To do:
        #unit test for (purposeful) fail cases....
        #clean-up the dumped text files.

    def test_zGetTol(self):
        print("\nTEST: zGetTol()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Try to set a valid tolerance operand
        self.link0.zSetTol(1,1,'TCON') # set tol operand of 1st row
        self.link0.zSetTol(1,2,1)      # set int1 =1
        self.link0.zSetTol(1,5,0.25)   # set min = 0.25
        self.link0.zSetTol(1,6,0.75)   # set max = 0.75
        tolData = self.link0.zGetTol(1)
        self.assertTupleEqual(tolData,('TCON', 1, 0, 0.25, 0.75, 0))

    def test_zGetTrace(self):
        print("\nTEST: zGetTrace()")
        # Load a lens file into the LDE (Not required to Push lens)
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Set up the data
        waveNum,mode,surf,hx,hy,px,py = 3,0,5,0.0,1.0,0.0,0.0
        rayTraceArg = (waveNum,mode,surf,hx,hy,px,py)
        expRayTraceData = (0, 0, 0.0, 2.750250805, 0.04747610066, 0.0,
                           0.2740755916, 0.9617081522, 0.0, 0.03451463936,
                           -0.9994041923, 1.0)
        # test returned tuple
        rayTraceData = self.link0.zGetTrace(*rayTraceArg)
        for i,d in enumerate(expRayTraceData):
            self.assertAlmostEqual(rayTraceData[i],d,places=4)
        (errorCode,vigCode,x,y,z,l,m,n,l2,
              m2,n2,intensity) = self.link0.zGetTrace(*rayTraceArg)
        traceDataTuple = (errorCode,vigCode,x,y,z,l,m,n,l2,m2,n2,intensity)
        for i,d in enumerate(expRayTraceData):
            self.assertAlmostEqual(traceDataTuple[i],d,places=4)
        # Check for individual values ... (not necessary)
        self.assertEqual(rayTraceData[0],errorCode)
        self.assertEqual(rayTraceData[1],vigCode)
        self.assertEqual(rayTraceData[2],x)
        self.assertEqual(rayTraceData[3],y)
        self.assertEqual(rayTraceData[4],z)
        self.assertEqual(rayTraceData[5],l)
        self.assertEqual(rayTraceData[6],m)
        self.assertEqual(rayTraceData[7],n)
        self.assertEqual(rayTraceData[8],l2)
        self.assertEqual(rayTraceData[9],m2)
        self.assertEqual(rayTraceData[10],n2)
        self.assertEqual(rayTraceData[11],intensity)
        if TestPyZDDEFunctions.pRetVar:
            print("Ray trace", rayTraceData)

    @unittest.skip("To implement")
    def test_zGetUDOSystem(self):
        print("\nTEST: zGetUDOSystem()")

    def test_zGetUpdate(self):
        print("\nTEST: zGetUpdate()")
        # Load & then push a lens file into the LDE
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Push the lens in the Zemax DDE server into the LDE
        ret = self.link0.zPushLens(updateFlag=1)
        # Update the lens to recompute
        ret = self.link0.zGetUpdate()
        self.assertIn(ret,(-998,-1,0))
        if ret == -1:
            print("MSG: ZEMAX couldn't update the lens")
        if ret == -998:
            print("MSG: zGetUpdate() function timed out")
        if TestPyZDDEFunctions.pRetVar:
            print("zGetUpdate return value", ret)


    def test_zGetVersion(self):
        print("\nTEST: zGetVersion()")
        ver = self.link0.zGetVersion()
        self.assertIsInstance(ver,int)
        if TestPyZDDEFunctions.pRetVar:
            print("VERSION: ", ver)

    def test_zGetWave(self):
        print("\nTEST: zGetWave()")
        # First set the waveslength data in the ZEMAX DDE server
        wavelength1 = 0.48613270
        wavelength2 = 0.58756180
        # set the number of wavelengths and primary wavelength
        waveData0 = self.link0.zSetWave(0,1,2)
        # set the wavelength data
        waveData1 = self.link0.zSetWave(1,wavelength1,0.5)
        waveData2 = self.link0.zSetWave(2,wavelength2,0.5)
        # Get the wavelength data using the zGetWave() function
        waveData_g0 = self.link0.zGetWave(0)
        waveData_g1 = self.link0.zGetWave(1)
        waveData_g2 = self.link0.zGetWave(2)
        if TestPyZDDEFunctions.pRetVar:
            print("Primary wavelength number = {}".format(waveData_g0[0]))
            print("Total number of wavelengths set = {}".format(waveData_g0[1]))
            print("Wavelength: {}, weight : {}".format(waveData_g1[0],waveData_g1[1]))
            print("Wavelength: {}, weight : {}".format(waveData_g2[0],waveData_g2[1]))

        #verify
        waveData_s_tuple = (waveData0[0],waveData0[1],waveData1[0],waveData1[1],
                                                      waveData2[0],waveData2[1],)
        waveData_g_tuple = (waveData_g0[0],waveData_g0[1],waveData_g1[0],waveData_g1[1],
                                                      waveData_g2[0],waveData_g2[1],)
        self.assertEqual(waveData_s_tuple,waveData_g_tuple)

    def test_zGetWaveTuple(self):
        print("\nTEST: zGetWaveTuple()")
        # First, set the wave fields in the ZEMAX DDE server
        # Create the wavelength and weight tuples
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        oWaveDataTuple_s = self.link0.zSetWaveTuple(iWaveDataTuple)
        # Now, call the zGetWaveTuple() to get teh wave data
        oWaveDataTuple_g = self.link0.zGetWaveTuple()
        if TestPyZDDEFunctions.pRetVar:
            print("Output wave data tuple",oWaveDataTuple_g)
        #verify that the returned wavelengths are same
        oWavelengths = oWaveDataTuple_g[0]
        for i,d in enumerate(oWavelengths):
            self.assertAlmostEqual(wavelengths[i],d,places=4)

    @unittest.skip("To implement")
    def test_zHammer(self):
        print("\nTEST: zHammer()")

    @unittest.skip("To implement")
    def test_zImportExtraData(self):
        print("\nTEST: zImportExtraData()")

    def test_zInsertConfig(self):
        print("\nTEST: zInsertConfig()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Get the current number of configurations (columns)
        currConfig = self.link0.zGetConfig()
        #Insert a config
        self.link0.zInsertConfig(currConfig[1]+1)
        #Assert if the number of configurations didn't increase, however the
        #current configuration shouldn't change, and the number of multiple
        #configurations must remain same.
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1]+1,currConfig[2]))

    def test_zInsertMCO(self):
        print("\nTEST: zInsertMCO()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Get the current number of configurations (columns and rows)
        currConfig = self.link0.zGetConfig()
        self.assertTupleEqual(currConfig,(1,1,1))
        #Insert a operand (row)
        newOperNumber = self.link0.zInsertMCO(2)
        self.assertEqual(newOperNumber,2)
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1],currConfig[2]+1))

    @unittest.skip("To implement")
    def test_zInsertObject(self):
        print("\nTEST: zInsertObject()")

    def test_zInsertSurface(self):
        print("\nTEST: zInsertSurface()")
        # Find the current number of surfaces
        systemData = self.link0.zGetSystem()
        init_surfaceNum = systemData[0]
        # call to insert surface
        retVal = self.link0.zInsertSurface(1)
        # verify that we now have the appropriate number of surfaces
        systemData = self.link0.zGetSystem()
        curr_surfaceNum = systemData[0]
        self.assertEqual(curr_surfaceNum,init_surfaceNum+1)

    def test_zLoadFile(self):
        print("\nTEST: zLoadFile()")
        global zmxfp
        #Try to load a non existant file
        filename = zmxfp+"nonExistantFile.zmx"
        ret = self.link0.zLoadFile(filename)
        self.assertEqual(ret,-999)
        #Now, try to load a real file
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        self.assertEqual(ret,0)
        if TestPyZDDEFunctions.pRetVar:
            print("zLoadFile return value:", ret)

    def test_zNewLens(self):
        print("\nTEST: zNewLens()")
        # Load the ZEMAX DDE server with a lens so that it has something to begin
        # with
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Call zNewLens to erase the current lens.
        ret = self.link0.zNewLens()
        if TestPyZDDEFunctions.pRetVar:
            print("zNewLens return val:", ret)
        #Call getSystem to see if we really have the "minimum" lens
        systemData = self.link0.zGetSystem()
        self.assertEqual(systemData[0],2,'numberOfSurfaces')
        self.assertEqual(systemData[1],0,'lens unit code')
        self.assertEqual(systemData[2],1,'stop surface-number')
        self.assertEqual(systemData[3],0,'non axial flag')
        self.assertEqual(systemData[4],0,'ray aiming type')
        self.assertEqual(systemData[5],0,'adjust index')
        self.assertEqual(systemData[6],20.0),'current temperature'
        self.assertEqual(systemData[7],1,'pressure')
        self.assertEqual(systemData[8],1,'global surface reference')
        #self.assertEqual(systemData[9],0,'need_save') #'need_save' deprecated

    def test_zPushLens(self):
        print("\nTEST: zPushLens()")
        # push a lens with an invalid flag, should rise ValueError exception
        self.assertRaises(ValueError,self.link0.zPushLens,updateFlag=10)
        # push a lens with valid flag.
        ret = self.link0.zPushLens()
        self.assertIn(ret,(0,-999,-998))
        ret = self.link0.zPushLens(updateFlag=1)
        self.assertIn(ret,(0,-999,-998))
        #Notify depending on return type
        # Note that the test as such should not "fail" if ZEMAX server returned
        # -999 (lens couldn't be pushed" or the function timed out (-998)
        if ret == -999:
            print("MSG: Lens could not be pushed into the LDE (check PushLensPermission)")
        if ret == -998:
            print("MSG: zPushLens() function timed out")
        if TestPyZDDEFunctions.pRetVar:
            "zPushLens return value:", ret

    def test_zPushLensPermission(self):
        print("\nTEST: zPushLensPermission()")
        status = self.link0.zPushLensPermission()
        self.assertIn(status,(0,1))
        if status:
            print("MSG: Push lens allowed")
        else:
            print("MSG: Push lens not allowed")
        if TestPyZDDEFunctions.pRetVar:
            print("zPushLens return status:", status)

    def test_zQuickFocus(self):
        print("\nTEST: zQuickFocus()")
        # Setup the system, wavelength, field points
        # Add some surfaces
        retVal = self.link0.zInsertSurface(1)
        #System
        unitCode,stopSurf,rayAimingType = 0,2,0  # mm, 4th,off
        useEnvData,temp,pressure,globalRefSurf = 0,20,1,1 # off, 20C,1ATM,ref=1st surf
        setSystemArg = (unitCode,stopSurf,rayAimingType,useEnvData,
                                                  temp,pressure,globalRefSurf)
        expSystemData = (2, 0, 2, 0, 0, 0, 20.0, 1, 1, 0)
        recSystemData = self.link0.zSetSystem(*setSystemArg)
        #Field
        iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                              # vdx=vdy=vcx=vcy=van=0.0
                           (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                           (0.0,10.0))                        # field3: xf=0.0,yf=10.0
        # Set the field data, such that fieldType is angle with rectangular normalization
        oFieldDataTuple = self.link0.zSetFieldTuple(0,1,iFieldDataTuple)
        # Wavelength
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        oWaveDataTuple = self.link0.zSetWaveTuple(iWaveDataTuple)
        # Now, call the zQuickFocus() function
        ret = self.link0.zQuickFocus(0,0) # RMS spot size, chief ray as reference
        print(ret)
        # I might need to have some suraces here.

    @unittest.skip("To implement test")
    def test_zSetAperture(self):
        print("\nTEST: zSetAperture()")
        pass

    def test_zSetConfig(self):
        print("\nTEST: zSetConfig()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        currConfig = self.link0.zGetConfig()
        #Since no configuration is initally present, it should return (1,1,1)
        self.assertTupleEqual(currConfig,(1,1,1))
        #Insert a config
        self.link0.zInsertConfig(currConfig[1]+1)
        #Assert if the number of configurations didn't increase, however the
        #current configuration shouldn't change, and the number of multiple
        #configurations must remain same.
        newConfig = self.link0.zGetConfig()
        self.assertTupleEqual(newConfig,(currConfig[0],currConfig[1]+1,currConfig[2]))
        #Now, finally, call zSetConfig() to switch configuration
        newConfig = self.link0.zSetConfig(2)
        self.assertEqual(newConfig[0],2)  # current configuration
        self.assertEqual(newConfig[1],2)  # number of configurations
        self.assertEqual(newConfig[2],0)  # error
        if TestPyZDDEFunctions.pRetVar:
            print("CONFIG: ", newConfig)
        #ToDo: Check error/test scenario

    @unittest.skip("To implement")
    def test_zSetExtra(self):
        print("\nTEST: zSetExtra()")


    def test_zSetField(self):
        print("\nTEST: zSetField()")
        # Set field with only 3 arguments, n=0
        # type = angle; 2 fields; rect normalization (default)
        fieldData = self.link0.zSetField(0,0,2)
        self.assertTupleEqual((0,2),(fieldData[0],fieldData[1]))
        # set field with 4 arguments, n=0
        fieldData = self.link0.zSetField(0,0,3,1)
        self.assertTupleEqual((0,3),(fieldData[0],fieldData[1]))
        #FIXME: zSetField is supposed to return more parameters.
        # is it a version issue?
        # set field with 3 args, n=1
        # 1st field, on-axis x, on-axis y, weight = 1 (default)
        fieldData = self.link0.zSetField(1,0,0)
        self.assertTupleEqual(fieldData,(0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0))
        # Set field with all input arguments, set first field
        fieldData = self.link0.zSetField(2,0,5,2.0,0.5,0.5,0.5,0.5,0.5)
        self.assertTupleEqual(fieldData,(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5))
        fieldData = self.link0.zSetField(3,0,10,1.0,0.0,0.0,0.0)
        self.assertTupleEqual(fieldData,(0.0, 10.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0))

    def test_zSetFieldTuple(self):
        print("\nTEST: zSetFieldTuple()")
        iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                              # vdx=vdy=vcx=vcy=van=0.0
                           (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                           (0.0,10.0))                        # field3: xf=0.0,yf=10.0
        # Set the field data, such that fieldType is angle with rectangular normalization
        oFieldDataTuple = self.link0.zSetFieldTuple(0,1,iFieldDataTuple)
        if TestPyZDDEFunctions.pRetVar:
            for i in range(len(iFieldDataTuple)):
                print("oFieldDataTuple, field {} : {}".format(i,oFieldDataTuple[i]))
        #Verify
        for i in range(len(iFieldDataTuple)):
            self.assertEqual(oFieldDataTuple[i][:len(iFieldDataTuple[i])],
                                                         iFieldDataTuple[i])

    @unittest.skip("To implement")
    def test_zSetFloat(self):
        print("\nTEST: zSetFloat()")
        pass

    @unittest.skip("To implement")
    def test_zSetLabel(self):
        print("\nTEST: zSetLabel()")
        pass

    def test_zSetMulticon(self):
        print("\nTEST: zSetMulticon()")
        #insert an additional configuration (column)
        self.link0.zInsertConfig(1)
        #insert an additional operand (row)
        self.link0.zInsertMCO(2)
        #Try to set invalid row operand at surface 2
        try:
            multiConData = self.link0.zSetMulticon(0,1,'INVALIDOPERAND',2,0,0)
        except ValueError:
            print("Expected Value Error raised")
        #Set the row operands (both to thickness, of surfaces 2, and 4 respectively)
        multiConData = self.link0.zSetMulticon(0,1,'THIC',2,0,0)
        self.assertTupleEqual(multiConData,('THIC',2,0,0))
        multiConData = self.link0.zSetMulticon(0,2,'THIC',4,0,0)
        self.assertTupleEqual(multiConData,('THIC',4,0,0))
        #Set configuration 1
        multiConData = self.link0.zSetMulticon(1,1,6.0076,0,1,1,1.0,0.0)
        self.assertTupleEqual(multiConData,(6.0076, 2, 2, 0, 1, 1, 1.0, 0.0))
        multiConData = self.link0.zSetMulticon(1,2,4.7504,0,1,1,1.0,0.0)
        self.assertTupleEqual(multiConData,(4.7504, 2, 2, 0, 1, 1, 1.0, 0.0))
        #Set configuration 2
        multiConData = self.link0.zSetMulticon(2,1,7.0000,0,1,1,1.0,0.0)
        self.assertTupleEqual(multiConData,(7.0, 2, 2, 0, 1, 1, 1.0, 0.0))
        multiConData = self.link0.zSetMulticon(2,2,5.0000,0,1,1,1.0,0.0)
        self.assertTupleEqual(multiConData,(5.0, 2, 2, 0, 1, 1, 1.0, 0.0))

    @unittest.skip("To implement test")
    def test_zSetNSCProperty(self):
        print("\nTEST: zSetNSCProperty()")
        pass

    def test_zSetNSCSetting(self):
        print("\nTEST: zSetNSCSetting()")
        pass

    def test_zSetPolState(self):
        print("\nTEST: zSetPolState()")
        #Set polarization of the "new" lens
        polStateData = self.link0.zSetPolState(0,0.5,0.5,10.0,10.0)
        self.assertTupleEqual(polStateData,(0,0.5,0.5,10.0,10.0))

    def test_zSetPrimaryWave(self):
        print("\nTEST: zSetPrimaryWave()")
        # first set 3 wavelength fields using zSetWaveTuple()
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        WaveDataTuple = self.link0.zSetWaveTuple(iWaveDataTuple)
        # right now, the first wavefield is the primary (0.48613270)
        # make the second wavelength field as the primary
        previousPrimary = self.link0.zGetWave(0)[0]
        primaryWaveNumber = 2
        oWaveData = self.link0.zSetPrimaryWave(primaryWaveNumber)
        if TestPyZDDEFunctions.pRetVar:
            print("Previous primary wavelength number =", previousPrimary)
            print("Current primary wavelength number =", oWaveData[0])
            print("Total number of wavelengths =", oWaveData[1])
        # verify
        self.assertEqual(primaryWaveNumber,oWaveData[0])
        self.assertEqual(len(wavelengths),oWaveData[1])

    @unittest.skip("To implement test")
    def test_zSetSolve(self):
        print("\nTEST: zSetSolve()")
        pass

    @unittest.skip("To implement")
    def test_zSetSurfaceData(self):
        print("\nTEST: zSetSurfaceData()")
        #Insert some surfaces

    @unittest.skip("To implement")
    def test_zSetSurfaceParameter(self):
        print("\nTEST: zSetSurfaceParameter()")
##        filename = zmxfp+lensFileName
##        ret = self.link0.zLoadFile(filename)
##        assert ret == 0
##        surfParam1 = self.link0.zGetSurfaceParameter(1,1)
##        print "Surface name: ", surfParam1
##        surfParam3 = self.link0.zGetSurfaceParameter(1,3)
##        print "Radius: ", surfParam3

    def test_zSetSystem(self):
        print("\nTEST: zSetSystem()")
        #Setup the arguments
        unitCode,stopSurf,rayAimingType = 0,4,0  # mm, 4th,off
        useEnvData,temp,pressure,globalRefSurf = 0,20,1,1 # off, 20C,1ATM,ref=1st surf
        setSystemArg = (unitCode,stopSurf,rayAimingType,useEnvData,
                                                  temp,pressure,globalRefSurf)
        expSystemData = (2, 0, 2, 0, 0, 0, 20.0, 1, 1)
        recSystemData = self.link0.zSetSystem(*setSystemArg)
        self.assertTupleEqual(expSystemData,recSystemData)
        if TestPyZDDEFunctions.pRetVar:
            systemDataPar = ('numberOfSurfaces','lens unit code',
                             'stop surface-number','non axial flag',
                             'ray aiming type','adjust index','current temperature',
                             'pressure','global surface reference') #'need_save' deprecated
            print("System data:")
            for i,elem in enumerate(systemDataPar):
                print("{} : {}".format(elem,recSystemData[i]))

    def test_zSetSystemAper(self):
        print("\nTEST: zSetSystemAper():")
        systemAperData_s = self.link0.zSetSystemAper(0,1,25) #sysAper=25mm,EPD
        self.assertEqual(systemAperData_s[0], 0, 'aperType = EPD')
        self.assertEqual(systemAperData_s[1], 1, 'stop surface number')
        self.assertEqual(systemAperData_s[2],25,'EPD value = 25 mm')

    def test_zSetSystemProperty(self):
        print("\nTEST: zSetSystemProperty():")
        #Set Aperture type as EPD
        sysPropData = self.link0.zSetSystemProperty(10,0)
        self.assertEqual(sysPropData,0)
        #Let lens title
        sysPropData = self.link0.zSetSystemProperty(16,"My Lens")
        self.assertEqual(sysPropData,"My Lens")
        #Set glass catalog
        sysPropData = self.link0.zSetSystemProperty(23,"SCHOTT HOYA OHARA")
        self.assertEqual(sysPropData,"SCHOTT HOYA OHARA")

    def test_zSetTol(self):
        print("\nTEST: zSetTol()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Try to set a wrong tolerance operand
        tolData = self.link0.zSetTol(1,1,'INVALIDOPERAND') # set tol operand of 1st row
        self.assertEqual(tolData,-1)
        #Try to set a valid tolerance operand
        tolData = self.link0.zSetTol(1,1,'TCON') # set tol operand of 1st row
        self.assertTupleEqual(tolData,('TCON', 0, 0, 0.0, 0.0, 0))

    def test_zSetTolRow(self):
        print("\nTEST: zSetTolRow()")
        # Load a lens file into the DDE server
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        #Try to set a wrong tolerance operand
        tolData = self.link0.zSetTolRow(1,'INVALIDOPERAND',1,0,0,0.25,0.75)
        self.assertEqual(tolData,-1)
        #Try to set a valid tolerance row data
        tolData = self.link0.zSetTolRow(1,'TRAD',1,0,0,0.25,0.75)
        self.assertTupleEqual(tolData,('TRAD', 1, 0, 0.25, 0.75, 0))

    @unittest.skip("To implement")
    def test_zSetUDOItem(self):
        print("\nTEST: zSetUDOItem()")

    def test_zSetWave(self):
        print("\nTEST: zSetWave()")
        wavelength1 = 0.48613270
        wavelength2 = 0.58756180
        # Call the zSetWave() function to set the primary wavelength & number
        # of wavelengths to set
        waveData = self.link0.zSetWave(0,1,2)
        if TestPyZDDEFunctions.pRetVar:
                print("Primary wavelength number = ", waveData[0])
                print("Total number of wavelengths set = ",waveData[1])
        # Verify
        self.assertEqual(waveData[0],1)
        self.assertEqual(waveData[1],2)
        # Set the first and second wavelength
        waveData1 = self.link0.zSetWave(1,wavelength1,0.5)
        waveData2 = self.link0.zSetWave(2,wavelength2,0.5)
        if TestPyZDDEFunctions.pRetVar:
            print("Wavelength: {}, weight: {}".format(waveData1[0],waveData1[1]))
            print("Wavelength: {}, weight: {}".format(waveData2[0],waveData2[1]))
        # Verify
        self.assertEqual(waveData1[0],wavelength1)
        self.assertEqual(waveData1[1],0.5)
        self.assertEqual(waveData2[0],wavelength2)
        self.assertEqual(waveData2[1],0.5)

    def test_zSetVig(self):
        print("\TEST: zSetVig()")
        retVal = self.link0.zSetVig()
        self.assertEqual(retVal,0)

    def test_zSetWaveTuple(self):
        print("\nTEST: zSetWaveTuple()")
        # Create the wavelength and weight tuples
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        oWaveDataTuple = self.link0.zSetWaveTuple(iWaveDataTuple)
        if TestPyZDDEFunctions.pRetVar:
            print("Output wave data tuple",oWaveDataTuple)
        #verify that the returned wavelengths are same
        oWavelengths = oWaveDataTuple[0]
        for i,d in enumerate(oWavelengths):
            self.assertAlmostEqual(wavelengths[i],d,places=4)

    @unittest.skip("Not necessary!")
    def test_zWindowMaximize(self):
        pass

    @unittest.skip("Not necessary!")
    def test_zWindowMinimize(self):
        pass

    @unittest.skip("Not necessary!")
    def test_zWindowRestore(self):
        pass

    @unittest.skip("Function not yet implemented")
    def test_zSetTimeout(self):
        print("\nTEST: zSetTimeout()")
        ret = self.link0.zSetTimeout(3)

if __name__ == '__main__':
    unittest.main()