#-------------------------------------------------------------------------------
# Name:        pyZDDEunittest.py
# Purpose:     pyZDDE unit test using the python unittest framework
#
# Author:      Indranil Sinharoy
#
# Created:     19/10/2012
# Copyright:   (c) Indranil 2012 - 2013
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision: 0.2
#-------------------------------------------------------------------------------

import os, sys, unittest

# Put both the "Test" and the "PyZDDE" directory in the python search
# path.
testdirectory = os.getcwd()
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
            print "Status for link 0:", status
        self.assertEqual(status,0)
        if status:
            print "Couln't initialize DDE."
        #Make sure to reset the lens
        ret = self.link0.zNewLens()

    def tearDown(self):
        # Tear down unit test
        if self.link0.connection:
            self.link0.zDDEClose()
        else:
            print "Server was already terminated"

    @unittest.skip("Init is now being called in the setup")
    def test_zDDEInit(self):
        # Test initialization
        print "\nTEST: zDDEInit()"
        status = self.link0.zDDEInit()
        print "Status for link 0:", status
        self.assertEqual(status,0)
        if status ==0:
            TestPyZDDEFunctions.b_zDDEInit_test_done = True
        else:
            print "Couln't initialize DDE."

    @unittest.skip("Being called in the teardown")
    def test_zDDEClose(self):
        print "\nTEST: zDDEClose()"
        stat = self.link0.zDDEClose()

    @unittest.skip("To implement")
    def test_zGetAperture(self):
        print "\nTEST: zGetAperture()"
        pass

    def test_zGetDate(self):
        print "\nTEST: zGetDate()"
        date = self.link0.zGetDate().rstrip()
        self.assertIsInstance(date,str)
        if TestPyZDDEFunctions.pRetVar:
            print "DATE: ", date

    def test_zGetField(self):
        print "\nTEST: zGetField()"
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
        print "\nTEST: zGetFieldTuple()"
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
                print "oFieldDataTuple_g, field",i,":",oFieldDataTuple_g[i]
        #Verify
        for i in range(len(iFieldDataTuple)):
            self.assertEqual(oFieldDataTuple_g[i][:len(iFieldDataTuple[i])],
                                                           iFieldDataTuple[i])

    def test_zGetFile(self):
        print "\nTEST: zGetFile()"
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0   # This is not a unit-test assert.
        reply = self.link0.zGetFile()
        self.assertEqual(reply,filename)
        if TestPyZDDEFunctions.pRetVar:
            print "zGetFile return value:", reply

    def test_zGetFirst(self):
        print "\nTEST: zGetFirst()"
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
            print "zGetFirst return value: ", focal,pwfn, rwfn,pima, pmag


    def test_zGetPupil(self):
        print "\nTEST: zGetPupil()"
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
            print "Pupil data:"
            print pupil_data[0],":",pupil_type[pupilData[0]]
            print pupil_data[1],":",pupilData[1],(pupil_value_type[0]
                                     if pupilData[0]==3 else pupil_value_type[1])
            for i in range(2,6):
                print pupil_data[i],":",pupilData[i]
            print pupil_data[6],":",apodization_type[pupilData[6]]
            print pupil_data[7],":",pupilData[7]

    def test_zGetRefresh(self):
        print "\nTEST: zGetRefresh()"
        # Load & push a lens file into the LDE
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        ret = self.link0.zPushLens(1)
        # Copy the lens data from the LDE into the stored copy of the ZEMAX
        # server.
        ret = self.link0.zGetRefresh()
        self.assertIn(ret,(-998,-1,0))
        if ret == -1:
            print "MSG: ZEMAX couldn't copy the lens data to the LDE"
        if ret == -998:
            print "MSG: zGetRefresh() function timed out"
        if TestPyZDDEFunctions.pRetVar:
            print "zGetRefresh return value", ret

    def test_zGetSerial(self):
        print "\nTEST: zGetSerial()"
        ser = self.link0.zGetSerial()
        self.assertIsInstance(ser,int)
        if TestPyZDDEFunctions.pRetVar:
            print "SERIAL:", ser

#    @unittest.skip("To implement")
    def test_zGetSurfaceData(self):
        print "\nTEST: zGetSurfaceData()"
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
            print "surfName :", surfName
            print "radius :", radius
            print "thickness: ", thick

    @unittest.skip("To implement")
    def test_zGetSurfaceDLL(self):
        print "\nTEST: zGetSurfaceDLL()"
        #Load a lens file

    @unittest.skip("To implement")
    def test_zGetSurfaceParameter(self):
        print "\nTEST: zGetSurfaceParameter()"
        #Load a lens file
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        assert ret == 0
        surfParam1 = self.link0.zGetSurfaceParameter(1,1)
        print "Surface name: ", surfParam1
        surfParam3 = self.link0.zGetSurfaceParameter(1,3)
        print "Radius: ", surfParam3


    def test_zGetSystem(self):
        print "\nTEST: zGetSystem()"
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
                             'pressure','global surface reference')  #'need_save' has been Deprecated.
            print "System data:"
            for i,elem in enumerate(systemDataPar):
                print elem,':',recSystemData_g[i]

    def test_zGetSystemApr(self):
        print "\nTEST: zGetSystemApr()"
        # First set the system aperture to known parameters in the ZEMAX server
        systemAperData_s = self.link0.zSetSystemAper(0,1,25) #sysAper=25mm,EPD
        systemAperData_g = self.link0.zGetSystemAper()
        self.assertTupleEqual(systemAperData_s,systemAperData_g)

    def test_zGetTextFile(self):
        print "\nTEST: zGetTextFile()"
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
            print "MSG: zGetTextFile failed"
        if ret == -998:
            print "MSG: zGetTextFile() function timed out"
        if TestPyZDDEFunctions.pRetVar:
            print "zGetTextFile return value", ret

        #Request zemax to dump prescription file, with a settings
        ret = self.link0.zGetRefresh()
        settingsFileName = "Cooke 40 degree field_PreSettings_OnlyCardinals.CFG"
        preFileName = 'Prescription_unitTest_01.txt'
        textFileName = testdirectory + '\\' + preFileName
        ret = self.link0.zGetTextFile(textFileName,'Pre',settingsFileName,1)
        self.assertIn(ret,(0,-1,-998)) #ensure that the ret is any valid return
        if ret == -1:
            print "MSG: zGetText failed"
        if ret == -998:
            print "MSG: zGetText() function timed out"
        if TestPyZDDEFunctions.pRetVar:
            print "zGetText return value", ret
        #To do:
        #unit test for (purposeful) fail cases....


    def test_zGetTrace(self):
        print "\nTEST: zGetTrace()"
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
            print "Ray trace", rayTraceData


    def test_zGetUpdate(self):
        print "\nTEST: zGetUpdate()"
        # Load & push a lens file into the LDE
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Push the lens in the Zemax DDE server into the LDE
        ret = self.link0.zPushLens(updateFlag=1)
        # Update the lens to recompute
        ret = self.link0.zGetUpdate()
        self.assertIn(ret,(-998,-1,0))
        if ret == -1:
            print "MSG: ZEMAX couldn't update the lens"
        if ret == -998:
            print "MSG: zGetUpdate() function timed out"
        if TestPyZDDEFunctions.pRetVar:
            print "zGetUpdate return value", ret


    def test_zGetVersion(self):
        print "\nTEST: zGetVersion()"
        ver = self.link0.zGetVersion()
        self.assertIsInstance(ver,int)
        if TestPyZDDEFunctions.pRetVar:
            print "VERSION: ", ver

    def test_zGetWave(self):
        print "\nTEST: zGetWave()"
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
            print "Primary wavelength number = ", waveData_g0[0]
            print "Total number of wavelengths set = ",waveData_g0[1]
            print "Wavelength: ",waveData_g1[0],
            print "weight", waveData_g1[1]
            print "Wavelength: ",waveData_g2[0],
            print "weight", waveData_g2[1]
        #verify
        waveData_s_tuple = (waveData0[0],waveData0[1],waveData1[0],waveData1[1],
                                                      waveData2[0],waveData2[1],)
        waveData_g_tuple = (waveData_g0[0],waveData_g0[1],waveData_g1[0],waveData_g1[1],
                                                      waveData_g2[0],waveData_g2[1],)
        self.assertEqual(waveData_s_tuple,waveData_g_tuple)

    def test_zGetWaveTuple(self):
        print "\nTEST: zGetWaveTuple()"
        # First, se the wave fields in the ZEMAX DDE server
        # Create the wavelength and weight tuples
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        oWaveDataTuple_s = self.link0.zSetWaveTuple(iWaveDataTuple)
        # Now, call the zGetWaveTuple() to get teh wave data
        oWaveDataTuple_g = self.link0.zGetWaveTuple()
        if TestPyZDDEFunctions.pRetVar:
            print "Output wave data tuple",oWaveDataTuple_g
        #verify
        self.assertTupleEqual(iWaveDataTuple,oWaveDataTuple_g)

    def test_zInsertSurface(self):
        print "\nTEST: zInsertSurface()"
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
        print "\nTEST: zLoadFile()"
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
            print "zLoadFile return value:", ret

    def test_zNewLens(self):
        print "\nTEST: zNewLens()"
        # Load the ZEMAX DDE server with a lens so that it has something to begin
        # with
        global zmxfp
        filename = zmxfp+lensFileName
        ret = self.link0.zLoadFile(filename)
        # Call zNewLens to erase the current lens.
        ret = self.link0.zNewLens()
        if TestPyZDDEFunctions.pRetVar:
            print "zNewLens return val:", ret
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
        #self.assertEqual(systemData[9],0,'need_save') #'need_save' has been deprecated

    def test_zPushLens(self):
        print "\nTEST: zPushLens()"
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
            print "MSG: Lens could not be pushed into the LDE (check PushLensPermission)"
        if ret == -998:
            print "MSG: zPushLens() function timed out"
        if TestPyZDDEFunctions.pRetVar:
            "zPushLens return value:", ret

    def test_zPushLensPermission(self):
        print "\nTEST: zPushLensPermission()"
        status = self.link0.zPushLensPermission()
        self.assertIn(status,(0,1))
        if status:
            print "MSG: Push lens allowed"
        else:
            print "MSG: Push lens not allowed"
        if TestPyZDDEFunctions.pRetVar:
            print "zPushLens return status:", status

    @unittest.skip("Issues")
    def test_zQuickFocus(self):
        print "\nTEST: zQuickFocus()"
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
        print ret
        # I might need to have some suraces here.



    @unittest.skip("To implement test")
    def test_zSetAperture(self):
        print "\nTEST: zSetAperture()"
        pass

    def test_zSetField(self):
        print "\nTEST: zSetField()"
        # Set field with only 3 arguments, n=0
        # type = angle; 2 fields; rect normalization (default)
        fieldData = self.link0.zSetField(0,0,2)
        self.assertTupleEqual((0,2),(fieldData[0],fieldData[1]))
        # set field with 4 arguments, n=0
        fieldData = self.link0.zSetField(0,0,3,1)
        self.assertTupleEqual((0,3),(fieldData[0],fieldData[1]))
        #ToDo: setfield is supposed to return more parameters.
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
        print "\nTEST: zSetFieldTuple()"
        iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                              # vdx=vdy=vcx=vcy=van=0.0
                           (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                           (0.0,10.0))                        # field3: xf=0.0,yf=10.0
        # Set the field data, such that fieldType is angle with rectangular normalization
        oFieldDataTuple = self.link0.zSetFieldTuple(0,1,iFieldDataTuple)
        if TestPyZDDEFunctions.pRetVar:
            for i in range(len(iFieldDataTuple)):
                print "oFieldDataTuple, field",i,":",oFieldDataTuple[i]
        #Verify
        for i in range(len(iFieldDataTuple)):
            self.assertEqual(oFieldDataTuple[i][:len(iFieldDataTuple[i])],
                                                           iFieldDataTuple[i])

    def test_zSetPrimary(self):
        print "\nTEST: zSetPrimary()"
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
            print "Previous primary wavelength number =", previousPrimary
            print "Current primary wavelength number =", oWaveData[0]
            print "Total number of wavelengths =", oWaveData[1]
        # verify
        self.assertEqual(primaryWaveNumber,oWaveData[0])
        self.assertEqual(len(wavelengths),oWaveData[1])

    @unittest.skip("To implement")
    def test_zSetSurfaceData(self):
        print "\nTEST: zSetSurfaceData()"
        #Insert some surfaces

    @unittest.skip("To implement")
    def test_zSetSurfaceParameter(self):
        print "\nTEST: zSetSurfaceParameter()"
##        filename = zmxfp+lensFileName
##        ret = self.link0.zLoadFile(filename)
##        assert ret == 0
##        surfParam1 = self.link0.zGetSurfaceParameter(1,1)
##        print "Surface name: ", surfParam1
##        surfParam3 = self.link0.zGetSurfaceParameter(1,3)
##        print "Radius: ", surfParam3

    def test_zSetSystem(self):
        print "\nTEST: zSetSystem()"
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
                             'pressure','global surface reference') #'need_save' has been deprecated
            print "System data:"
            for i,elem in enumerate(systemDataPar):
                print elem,':',recSystemData[i]

    def test_zSetSystemAper(self):
        print "\nTEST: zSetSystemAper():"
        systemAperData_s = self.link0.zSetSystemAper(0,1,25) #sysAper=25mm,EPD
        self.assertEqual(systemAperData_s[0], 0, 'aperType = EPD')
        self.assertEqual(systemAperData_s[1], 1, 'stop surface number')
        self.assertEqual(systemAperData_s[2],25,'EPD value = 25 mm')

    def test_zSetWave(self):
        print "\nTEST: zSetWave()"
        wavelength1 = 0.48613270
        wavelength2 = 0.58756180
        # Call the zSetWave() function to set the primary wavelength & number
        # of wavelengths to set
        waveData = self.link0.zSetWave(0,1,2)
        if TestPyZDDEFunctions.pRetVar:
                print "Primary wavelength number = ", waveData[0]
                print "Total number of wavelengths set = ",waveData[1]
        # Verify
        self.assertEqual(waveData[0],1)
        self.assertEqual(waveData[1],2)
        # Set the first and second wavelength
        waveData1 = self.link0.zSetWave(1,wavelength1,0.5)
        waveData2 = self.link0.zSetWave(2,wavelength2,0.5)
        if TestPyZDDEFunctions.pRetVar:
            print "Wavelength: ",waveData1[0],
            print "weight", waveData1[1]
            print "Wavelength: ",waveData2[0],
            print "weight", waveData2[1]
        # Verify
        self.assertEqual(waveData1[0],wavelength1)
        self.assertEqual(waveData1[1],0.5)
        self.assertEqual(waveData2[0],wavelength2)
        self.assertEqual(waveData2[1],0.5)

    def test_zSetWaveTuple(self):
        print "\nTEST: zSetWaveTuple()"
        # Create the wavelength and weight tuples
        wavelengths = (0.48613270,0.58756180,0.65627250)
        weights = (1.0,1.0,1.0)
        iWaveDataTuple = (wavelengths,weights)
        oWaveDataTuple = self.link0.zSetWaveTuple(iWaveDataTuple)
        if TestPyZDDEFunctions.pRetVar:
            print "Output wave data tuple",oWaveDataTuple
        #verify
        self.assertTupleEqual(oWaveDataTuple,iWaveDataTuple)


    @unittest.skip("Function not yet implemented")
    def test_zSetTimeout(self):
        print "\nTEST: zSetTimeout()"
        ret = self.link0.zSetTimeout(3)



if __name__ == '__main__':
    unittest.main()


#ToDo:
# 1. all prints with "MSG" should be put into a single buffer and printed
#    as a summery at the very end of the complete test. there are some facilities
#    for doing this @ unittest framework (TestResult object??)
# 2. zGetText() create more scenarios in the unit test of this function
#    also, the unit test should read the "dumped" file and then verify the expected
#    data. Then before exiting the test, the "dumped" files should be deleted.


# To do next:
# 1. zSetField (done)
# 2. zGetField (done)
# 5. zGetSolve
# 6. zSetSolve
# 7. zQuickFocus (done, to test)
# 8. zSetSurfaceParameter
# 9. zGetSurfaceParameter
#10. zSetFieldMatrix (MZDDE functions) --> zSetFieldTuple (done)
#11. zGetFieldMatrix (MZDDE functions) --> zGetFieldTuple (done)
#12. zSetAperture (done, to test)
#13. zGetAperture (done, to test)


# To do in near future:
# 1. A function similar to "help zemaxbuttons" implemented in MZDDE. Useful when
#    someone will quickly want to know the 3 letter codes for the different functions
#    especially when using functions like zGetText( ).
# 2. The test_pyZDDE() is not functional at this point in time in the latest version
#    of Zemax. It used to be functional in older version of Zemax (and it still does).
#    Need to fix it.
# 3. Add one more case to test_zPushLens() .... load a lens into the DDE server and
#    then push it to the LDE

# To check:
# It seems that the zGetField(0) and the return of zSetField(0) is returning only
# 2 arguments (type,numberoffields) when it is expected to return 5
# (type,numberoffields,x_field_max,y_field_max, normalization) ... why is this
# behavior? is it because I am testing it in an older version of ZEMAX? if that
# turns out to be the case, then can you use "version number" to do conditional
# tests?

# when the ZEMAX DDE server returns multiple values within a string, it can
# contain the characters '\r\n' such as '5.000000000E-001\r\n' or ['0', '3\r\n']
# Usually, it is not a problem as the "\r\n" parts are automatically stripped
# when a type converstion from string to int or float is done; so an extra
# regex is not necessary. If the returned "string" contains "\r\n' in the end,
# then just use reply.rstrip() in the main function (not in the unit-test function)
#to get rid of it.
