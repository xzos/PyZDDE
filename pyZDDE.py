#-------------------------------------------------------------------------------
# Name:        pyZDDE.py
# Purpose:     Python based DDE link with ZEMAX server, similar to Matlab based
#              MZDDE toolbox.
# Author:      Indranil Sinharoy
#
# Created:     08/10/2012
# Copyright:   (c)  2012 - 2013
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.2
#-------------------------------------------------------------------------------

import win32ui
import dde
import os
from os import path
from sys import exc_info
from math import pi,cos,sin
import warnings

DEBUG_PRINT_LEVEL = 0 # 0=No debug prints, but allow all essential prints
                      # 1 to 2 levels of debug print, 2 = print all

# Helper functions
def debugPrint(level,msg):
    """
    args:
        level = 0, message will definitely be printed
                1 or 2, message will be printed if level >= DEBUG_PRINT_LEVEL
        msg = string message to pint
    """
    global DEBUG_PRINT_LEVEL
    if level <= DEBUG_PRINT_LEVEL:
        print "DEBUG PRINT (Level" + str(level)+ ":)" + msg
    return

class pyzdde(object):
    __chNum = -1          # channel Number
    __liveCh = 0          # number of live channels.
    __server = 0
    DDE_TIMEOUT = 3000    # Not implemented (for future), timeout (pywin32 DDE = 1 min)

    def __init__(self):
        pyzdde.__chNum +=1   # increment ch. count when DDE ch. is instantiated.
        self.appName = "ZEMAX"+str(pyzdde.__chNum) if pyzdde.__chNum > 0 else "ZEMAX"
        self.connection = False  # 1/0 depending on successful connection or not

    # ZEMAX <--> PyZDDE client connection methods
    #--------------------------------------------
    def zDDEInit(self):
        """Initiates DDE link with Zemax server.

        zDDEInit( ) -> status

        status:
            0 : DDE link to ZEMAX was successfully established.
           -1 : DDE link couldn't be established.

        The function also sets the timeout value for all ZEMAX DDE calls to 3 sec.
        The timeout is not implemented now.

        See also zDDEClose, zDDEStart, zSetTimeout
        """
        debugPrint(1,"appName = " + self.appName)
        debugPrint(1,"liveCh = " + str(pyzdde.__liveCh))
        if self.appName=="ZEMAX" or pyzdde.__liveCh==0: # do this only one time or when there is no channel
            pyzdde.server = dde.CreateServer()
            pyzdde.server.Create("ZCLIENT")           # Name of the client
        # Try to create individual conversations for each ZEMAX application.
        self.conversation = dde.CreateConversation(pyzdde.server)
        try:
            self.conversation.ConnectTo(self.appName," ")
        except:
            info = exc_info()
            print "Error:", info[1],
            print ": ZEMAX may not have been started"
            return -1
        else:
            debugPrint(1,"Zemax instance successfully connected")
            pyzdde.__liveCh += 1 # increment the number of live channels
            self.connection = True
            #DDE_TIMEOUT = 3000 #The default timeout(FIXME: Not yet implemented).
            return 0

    def zDDEClose(self):
        """Close the DDE link with Zemax server.

        zDDEClose( ) -> Status

        Status = 0 on success.
        """
        # Close the server only if a channel was truely established and it is the last one.
        if self.connection and pyzdde.__liveCh <=1:
            self.server.Shutdown()
            self.connection = False
            pyzdde.__liveCh -=1  # This will become zero now. (reset)
            pyzdde.__chNum = -1  # Reset the chNum ...
            debugPrint(2,"server shutdown")
        elif self.connection:  # if additional channels were successfully created.
            self.connection = False
            pyzdde.__liveCh -=1
            debugPrint(2,"liveCh decremented without shutting down DDE channel")
        else:   # if zDDEClose is called by an object which didn't have a channel anyways
            debugPrint(3,"Nothing to do")

        return 0              # For future compatibility

    def zSetTimeout(self,time):
        """ sets the timeout in seconds for all ZEMAX DDE calls.

        zSetTimeOut(time)
        args:
            time: time in seconds.

        See also zDDEInit, zDDEStart
        """
        warnings.warn("Not implemented. Default timeout = 1 min")
        pyzdde.DDE_TIMEOUT = round(time*1000) # set time in milliseconds

    def __del__(self):
        """Destructor"""
        debugPrint(3,"Destructor called")
        self.zDDEClose()

    # ZEMAX control/query methods
    #----------------------------
    def zDeleteConfig(self,number):
        """Deletes an existing configuration (column) in the multi-configuration
        editor.

        zDeleteConfig(config)->configNumber

        args:
            number : (integer) configuration number to delete
        ret:
            retVal : (integer) configuration number deleted.

        Note: After deleting the configuration, all succeeding configurations are
        re-numbered.

        See also zInsertConfig.
        """
        return int(self.conversation.Request("DeleteConfig,"+str(number)))

    def zDeleteMCO(self,operandNumber):
        """Deletes an existing operand (row) in the multi-configuration editor.

        zDeleteMCO(operandNumber)->retVal

        args:
            operandNumber : (integer) operand number (row in the MCE) to delete.
        ret:
            retVal        : (integer) new number of operands.

        Note: After deleting the row, all succeeding rows (operands) are
        re-numbered.

        See also zInsertMCO.
        """
        return int(self.conversation.Request("DeleteMCO,"+str(operandNumber)))

    def zGetAperture(self,surfNum):
        """Get the surface aperture data.

        zGetAperture(surfNum) -> apertureInfo

        args:
            surfNum : surface number

        ret:
            aType        : integer code to specify aperture type
                            0 = no aperture (na)
                            1 = circular aperture (ca)
                            2 = circular obscuration (co)
                            3 = spider (s)
                            4 = rectangular aperture (ra)
                            5 = rectangular obscuration (ro)
                            6 = elliptical aperture (ea)
                            7 = elliptical obscuration (eo)
                            8 = user defined aperture (uda)
                            9 = user defined obscuration (udo)
                            10 = floating aperture (fa)
            aMin         : min radius(ca), min radius(co),width of arm(s),
                           X-half width(ra),X-half width(ro),X-half width(ea),
                           X-half width(eo)
            aMax         : max radius(ca), max radius(co),number of arm(s),
                           X-half width(ra),X-half width(ro),X-half width(ea),
                           X-half width(eo)
                           see "Aperture type and other aperture controls" for more details.
            xDecenter    : amount of decenter from current optical axis (lens units)
            yDecenter    : amount of decenter from current optical axis (lens units)
            apertureFile : a text file with .UDA extention. see "User defined
                           apertures and obscurations" in ZEMAX manual for more details.

        """
        reply = self.conversation.Request("GetAperture,"+str(surfNum))
        rs = reply.split(',')
        apertureInfo = [int(rs[i]) if i==5 else float(rs[i])
                                             for i in range(len(rs[:-1]))]
        apertureInfo.append(rs[-1].rstrip()) #append the test file (string)
        return tuple(apertureInfo)

    def zGetConfig(self):
        """Returns the current configuration number, the number of configurations,
        and the number of multiple configuration operands.

        zGetConfig()->(currentConfig, numberOfConfigs, numberOfMutiConfigOper)

        args:
            none
        ret:
            3-tuple containing currentConfig, numberOfConfigs, numberOfMutiConfigOper

        Note:
            The function returns (1,1,1) even if the multi-configuration editor
            is empty. This is because, by default, the current lens in the LDE
            is, by default, set to the current configuration. The initial number
            of configurations is therefore 1, and the number of operators in the
            multi-configuration editor is also 1 (generally, MOFF).

        See also zSetConfig. Use zInsertConfig to insert new configuration in the
        multi-configuration editor.
        """
        reply = self.conversation.Request('GetConfig')
        rs = reply.split(',')
        return tuple([int(elem) for elem in rs])

    def zGetDate(self):
        """Request current date from the ZEMAX DDE server.

        zGetDate()->date

        ret:
            date: date is a string.
        """
        return self.conversation.Request('GetDate')

    def zGetExtra(self,surfaceNumber,columnNumber):
        """Returns extra surface data from the Extra Data Editor

        zGetExtra(surfaceNumber,columnNumber)->value

        arg:
            surfaceNumber : (integer) surface number
            columnNumber  : (integer) column number
        ret:
            value         : (float) numeric data value

        See also zSetExtra
        """
        cmd = ('GetExtra,%i,%i')%(surfaceNumber,columnNumber)
        reply = self.conversation.Request(cmd)
        return float(reply)

    def zGetField(self,n):
        """Extract field data from ZEMAX DDE server

        zGetField(n) -> fieldData

        args [if n =0]:
            n: for n=0, the function returns general field parameters.

        args [if 0 < n <= number of fields]:
            n: field number

        ret [if n=0]: fieldData is a tuple containing the following
            type                 : integer (0=angles in degrees, 1=object height
                                            2=paraxial image height,
                                            3=real image height)
            number               : number of fields currently defined
            max_x_field          : values used to normalize x field coordinate
            max_y_field          : values used to normalize y field coordinate
            normalization_method : field normalization method (0=radial, 1 = rectangular)

        ret [if 0 < n <= number of fields]: fieldData is a tuple containing the following
            xf     : the field x value
            yf     : the field y value
            wgt    : field weight
            vdx    : decenter x vignetting factor
            vdy    : decenter y vignetting factor
            vcx    : compression x vignetting factor
            vcy    : compression y vignetting factor
            van    : angle vignetting factor

        Note: the returned tuple's content and structure is exactly same as that
        of zSetField()

        See also zSetField()
        """
        reply = self.conversation.Request('GetField,'+str(n))
        rs = reply.split(',')
        if n: # n > 0
            fieldData = tuple([float(elem) for elem in rs])
        else: # n = 0
            fieldData = tuple([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return fieldData

    def zGetFieldTuple(self):
        """Get all field data in a single N-D tuple.

        zGetFieldTuple()->fieldDataTuple

        ret:
            fieldDataTuple: the output field data tuple is also a N-D tuple (0<N<=12)
            with every dimension representing a single field location. Each
            dimension has all 8 field parameters.

        See also zGetField(), zSetField(), zSetFieldTuple()
        """
        fieldCount = self.zGetField(0)[1]
        fieldDataTuple = [ ]
        for i in range(fieldCount):
            reply = self.conversation.Request('GetField,'+str(i+1))
            rs = reply.split(',')
            fieldData = tuple([float(elem) for elem in rs])
            fieldDataTuple.append(fieldData)
        return tuple(fieldDataTuple)

    def zGetFile(self):
        """This method extracts and returns the full name of the lens, including
           the drive and path.

           zGetFile()-> file_name

            ret:
                file_name: filename of the Zemax file that is currently present
                in the Zemax DDE server.

           Note:
           1. Extreme caution should be used if the file is to be tampered with;
              since at any time ZEMAX may read or write from/to this file.
        """
        reply = self.conversation.Request('GetFile')
        return reply.rstrip()

    def zGetFirst(self):
        """Returns the first order data about the lens.

            zGetFirst()->(focal, pwfn, rwfn, pima, pmag)

            ret:
                The function returns a 5-tuple containing the following:
                focal   : the Effective Focal Length (EFL) in lens units,
                pwfn    : the paraxial working F/#,
                rwfn    : real working F/#,
                pima    : paraxial image height, and
                pmag    : paraxial magnification.
        """
        reply = self.conversation.Request('GetFirst')
        rs = reply.split(',')
        return tuple([float(elem) for elem in rs])

    def zGetMulticon(self,config,row):
        """Extract data from the multi-configuration editor.

        zGetMulticon(config,row)->multiConData

        args:
            config : (integer)
            row    :
        """
        reply = self.conversation.Request('GetMulticon,%i,%i'%(config,row))
        if config: # if config > 0
            rs = reply.split(",")
            print rs
            multiConData = [float(rs[i]) if (i == 0 or i == 6 or i== 7) else float(rs[i])
                                                 for i in range(len(rs))]
        else: # if config == 0
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0,rs[0])
        return tuple(multiConData)


    def zGetPupil(self):
        """Get pupil data from ZEMAX.

        zGetPupil()-> pupilData

        ret:
            pupilData: a tuple containing the following elements:
                aType              : integer indicating the system aperture
                                     0 = entrance pupil diameter
                                     1 = image space F/#
                                     2 = object space NA
                                     3 = float by stop
                                     4 = paraxial working F/#
                                     5 = object cone angle
                value              : if aperture type == float by stop
                                         value is stop surface semi-diameter
                                     else
                                         value is the sytem aperture
                ENPD               : entrance pupil diameter (in lens units)
                ENPP               : entrance pupil position (in lens units)
                EXPD               : exit pupil diameter (in lens units)
                EXPP               : exit pupil position (in lens units)
                apodization_type   : integer indicating the following types
                                     0 = none
                                     1 = Gaussian
                                     2 = Tangential/Cosine cubed
                apodization_factor : number shown on general data dialog box.

        """
        reply = self.conversation.Request('GetPupil')
        rs = reply.split(',')
        pupilData = tuple([int(elem) if (i==0 or i==6)
                                 else float(elem) for i,elem in enumerate(rs)])
        return pupilData

    def zGetRefresh(self):
        """Copy the lens data from the LDE into the stored copy of the ZEMAX
        server.The lens is then updated, and ZEMAX re-computes all data.

        zGetRefresh() -> status

        ret:
            status:    0 if successful,
                      -1 if ZEMAX could not copy the lens data from LDE to the server
                    -998 if the command times out (Note MZDDE returns -2)

        If zGetRefresh() returns -1, no ray tracing can be performed.

        See also zGetUpdate, zPushLens.
        """
        reply = None
        reply = self.conversation.Request('GetRefresh')
        if reply:
            return int(reply) #Note: Zemax returns -1 if GetRefresh fails.
        else:
            return -998

    def zGetSerial(self):
        """Get the serial number
        """
        reply = self.conversation.Request('GetSerial')
        return int(reply)

    def zGetSurfaceData(self,surfaceNumber,code,arg2=None):
        """Gets surface data on a sequential lens surface.

        zGetSurfaceData(surfaceNum,code,[, arg2])-> surfaceDatum

        args:
            surfaceNum : the surface number
            code       : integer number (see below)
            arg2       : (Optional) for item codes above 70.

        Gets surface datum at surfaceNumber depending on the code according to
        the following table.
        The code is as shown in the following table.
        arg2 is required for some item codes. [Codes above 70]


        Code      - Data returned by zGetSurfaceData()
        0         - Surface type name. (string)
        1         - Comment. (string)
        2         - Curvature (numeric). [Note: It is not radius!!]
        3         - Thickness. (numeric)
        4         - Glass. (string)
        5         - Semi-Diameter. (numeric)
        6         - Conic. (numeric)
        7         - Coating. (string)
        8         - Thermal Coefficient of Expansion (TCE).
        9         - User-defined .dll (string)
        20        - Ignore surface flag, 0 for not ignored, 1 for ignored.
        51        - Tilt, Decenter order before surface; 0 for Decenter then
                    Tilt, 1 for Tilt then Decenter.
        52        - Decenter x
        53        - Decenter y
        54        - Tilt x before surface
        55        - Tilt y before surface
        56        - Tilt z before surface
        60        - Status of Tilt/Decenter after surface. 0 for explicit, 1
                    for pickup current surface,2 for reverse current surface,
                    3 for pickup previous surface, 4 for reverse previous
                    surface,etc.
        61        - Tilt, Decenter order after surface; 0 for Decenter then
                    Tile, 1 for Tilt then Decenter.
        62        - Decenter x after surface
        63        - Decenter y after surface
        64        - Tilt x after surface
        65        - Tilt y after surface
        66        - Tilt z after surface
        70        - Use Layer Multipliers and Index Offsets. Use 1 for true,
                    0 for false.
        71        - Layer Multiplier value. The coating layer number is defined
                    by arg2.
        72        - Layer Multiplier status. Use 0 for fixed, 1 for variable,
                    or n+1 for pickup from layer n. The coating layer number
                    is defined by arg2.
        73        - Layer Index Offset value. The coating layer number is
                    defined by arg2.
        74        - Layer Index Offset status. Use 0 for fixed, 1 for variable,
                    or n+1 for pickup from layer n.The coating layer number is
                    defined by arg2.
        75        - Layer Extinction Offset value. The coating layer number is
                    defined by arg2.
        76        - Layer Extinction Offset status. Use 0 for fixed, 1 for
                    variable, or n+1 for pickup from layer n. The coating layer
                    number is defined by arg2.
        Other     - Reserved for future expansion of this feature.

        See also zSetSurfaceData, zGetSurfaceParameter and ZemaxSurfTypes
        """
        if arg2== None:
            cmd = ('GetSurfaceData,%i,%i')%(surfaceNumber,code)
        else:
            cmd = ('GetSurfaceData,%i,%i,%i')%(surfaceNumber,code,arg2)
        reply = self.conversation.Request(cmd)
        if code in (0,1,4,7,9):
            surfaceDatum = reply.split()[0]
        else:
            surfaceDatum = float(reply)
        return surfaceDatum

    def zGetSurfaceDLL(self,surfaceNumber):
        """Return the name of the DLL if the surface is a user defined type.

        zGetSurfaceDLL(surfaceNumber)->(dllName,surfaceName)

        args:
            surfaceNumber: (integer) surface number of the user defined surface
        ret:
            Returns a tuble with the following elements
            dllName      : (string) The name of the defining DLL
            surfaceName  : (string) surface name displayed by the DLL in the surface
                           type column of the LDE.

        """
        cmd = ('GetSurfaceDLL,%i')%(surfaceNumber)
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        return (rs[0],rs[1])

    def zGetSurfaceParameter(self,surfaceNumber,parameter):
        """Return the surface parameter data for the surface associated with the
        given surfaceNumber

        zGetSurfaceParameter(surfaceNumber,parameter)->parameterData

        args:
            surfaceNumber  : (integer) surface number of the surface
            parameter      : (integer) parameter (Par in LDE) number being queried
        ret:
            parameterData  : (float) the parameter value

        Note: To get thickness, radius, glass, semi-diameter, conic, etc, use
        zGetSurfaceData()
        See also zGetSurfaceData, ZSetSurfaceParameter.
        """
        cmd = ('GetSurfaceParameter,%i,%i')%(surfaceNumber,parameter)
        reply = self.conversation.Request(cmd)
        return float(reply)


    def zGetSystem(self):
        """Gets a number of general lens system data (General Lens Data)

        zGetSystem() -> systemData

        ret:
            systemData : the systemData is a tuple with the following elements:
                numSurfs      : number of surfaces
                unitCode      : lens units code (0,1,2,or 3 for mm, cm, in, or M)
                stopSurf      : the stop surface number
                nonAxialFlag  : flag to indicate if system is non-axial symmetric
                                (0 for axial, 1 if not axial)
                rayAimingType : ray aiming type (0,1, or 2 for off, paraxial or real)
                adjust_index  : adjust index data to environment flag (0 if false, 1 if true)
                temp          : the current temperature
                pressure      : the current pressure
                globalRefSurf : the global coordinate reference surface number
                need_save     : indicates whether the file has been modified. [Deprecated]

        Note: the returned data structure is exactly similar to the data structure
        returned by the zSetSystem() method.

        See also zSetSystem, zGetSystemAper, zGetAperture, zSetAperture
        """
        reply = self.conversation.Request("GetSystem")
        rs = reply.split(',')
        systemData = tuple([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zGetSystemAper(self):
        """Gets system aperture data.

        zGetSystemAper()-> systemAperData

        ret:
            systemAperData: systemAperData is a tuple containing the following
                aType              : integer indicating the system aperture
                                     0 = entrance pupil diameter
                                     1 = image space F/#
                                     2 = object space NA
                                     3 = float by stop
                                     4 = paraxial working F/#
                                     5 = object cone angle
                stopSurf           : stop surface
                value              : if aperture type == float by stop
                                         value is stop surface semi-diameter
                                     else
                                         value is the sytem aperture

        Note: the returned tuple is the same as the returned tuple of zSetSystemAper()

        See also, zGetSystem(), zSetSystemAper()
        """
        reply = self.conversation.Request("GetSystemAper")
        rs = reply.split(',')
        systemAperData = tuple([float(elem) for elem in rs])
        return systemAperData

    def zGetTextFile(self,textFileName, analysisType, settingsFileName, flag):
        """Request Zemax to save a text file for any analysis that supports text
           output.

           zGetText(textFilename, analysisType, settingsFilename, flag) -> retVal

           args:
            textFileName : name of the file to be created including the full path,
                           name, and extension for the text file.
            analysisType : 3 letter case-sensitive label that indicates the
                           type of the analysis to be performed. They are identical
                           to those used for the button bar in Zemax. The labels
                           are case sensitive. If no label is provided or recognized,
                           a standard raytrace will be generated. If a valid file
                           name is used for the "settingsFileName", ZEMAX will use
                           or save the settings used to compute the text file,
                           depending upon the value of the flag parameter.
            flag        :  0 = default settings used for the text
                           1 = settings provided in the settings file, if valid,
                               else default settings used
                           2 = settings provided in the settings file, if valid,
                               will be used and the settings box for the requested
                               feature will be displayed. After the user makes any
                               changes to the settings the text will then be
                               generated using the new settings.
                           Please see the ZEMAX manual for more details.
           retVal:
            0      : Success
            -1     : Text file could not be saved (Zemax may not have received
                     a full path name or extention).
            -998   : Command timed out

        Notes: No matter what the flag value is, if a valid file name is provided
        for the settingsfilename, the settings used will be written to the settings
        file, overwriting any data in the file.

        See also zGetMetaFile, zOpenWindow.
        """
        retVal = -1
        #Check if the file path is valid and has extension
        if path.isabs(textFileName) and path.splitext(textFileName)[1]!='':
            cmd = 'GetTextFile,"%s",%s,"%s",%i' %(textFileName,analysisType,
                                                  settingsFileName,flag)
            reply = self.conversation.Request(cmd)
            if reply.split()[0] == 'OK':
                retVal = 0
        return retVal


    def zGetTrace(self,waveNum,mode,surf,hx,hy,px,py):
        """Trace a ray through the current lens in the ZEMAX DDE server.

        zGetTrace(waveNum,mode,surf,hx,hy,px,py) -> rayTraceData

        args:
            waveNum : wavelength number as in the wavelength data editor
            mode    : 0 = real, 1 = paraxial
            surf    : surface to trace the ray to. Usually, the ray data is only
                      needed at the image surface; setting the surface number to
                      -1 will yield data at the image surface.
            hx      : normalized field height along x axis
            hy      : normalized field height along y axis
            px      : normalized height in pupil coordinate along x axis
            py      : normalized height in pupil coordinate along y axis
        ret:
            rayTraceData : rayTraceData is a tuple containing the following elements:
                errorCode : 0, if the ray traced successfully
                            +ve number indicates that the ray missed the surface
                            -ve number indicates that the ray total internal reflected
                            (TIR) at the surface given by the absolute value of the
                            errorCode number.
                vigCode   : the first surface where the ray was vignetted. Unless an
                            error occurs at that surface or subsequent to that surface,
                            the ray will continue to trace to the requested surface.
                x,y,z     : coordinates of the ray on the requested surface
                l,m,n     : the direction cosines after refraction into the media
                            following the requested surface.
                l2,m2,n2  : the surface intercept direction normals at the requested
                            surface.
                intensity : the relative transmitted intensity of the ray, including
                            any pupil or surface apodization defined.

        Example: To trace the real chief ray to surface 5 at wavelength 3, use
            rayTraceData = zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

            OR

            (errorCode,vigCode,x,y,z,l,m,n,l2,m2,n2,intensity) = \
                                              zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

        Note: Use of zGetTrace() has significant overhead as only one ray per DDE call
        is traced. Please refer to the ZEMAX manual for more details. Also, if a large
        number of rays are to be traced, see the section "Tracing large number of rays"
        in the ZEMAX manual.

        See also zGetTraceDirect, zGetPolTrace, zGetPolTraceDirect
        """
        cmd = 'GetTrace,%i,%i,%i,%1.4f,%1.4f,%1.4f,%1.4f' %(waveNum,mode,surf,
                                                                   hx,hy,px,py)
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        rayTraceData = tuple([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return rayTraceData

    def zGetUpdate(self):
        """Update the lens, which means Zemax recomputes
           all pupil positions, solves, and index data.

           zGetUpdate() -> status

            status :   0 = Zemax successfully updated the lens
                      -1 = No raytrace performed
                    -998 = Command timed

           To update the merit function, use the zOptimize item with the number
           of cycles set to -1.
           See also zGetRefresh, zOptimize, zPushLens
        """
        status,ret = -998, None
        ret = self.conversation.Request("GetUpdate")
        if ret != None:
            status = int(ret)  #Note: Zemax returns -1 if GetUpdate fails.
        return status

    def zGetVersion(self):
        """Get the current version of ZEMAX which is running.

        zGetVersion() -> version (integer, generally 5 digit)

        """
        return int(self.conversation.Request("GetVersion"))

    def zGetWave(self,n):
        """Extract wavelength data from ZEMAX DDE server.

        There are 2 ways of using this function:
            zGetWave(0)-> waveData
              OR
            zGetWave(wavelengthNumber)-> waveData

        ret:
            if n==0: waveData is a tuple containing the following:
                primary : number indicating the primary wavelength (integer)
                number  : number of wavelengths currently defined (integer).
            elif 0 < n <= number of wavelengths: waveData consists of:
                wavelength : value of the specific wavelength (floating point)
                weight     : weight of the specific wavelength (floating point)

        Note: the returned tuple is exactly same in structure and contents to that
        returned by zSetWave().

        See also zSetWave(),zSetWaveTuple(), zGetWaveTuple().
        """
        reply = self.conversation.Request('GetWave,'+str(n))
        rs = reply.split(',')
        if n:
            waveData = tuple([float(ele) for ele in rs])
        else:
            waveData = tuple([int(ele) for ele in rs])
        return waveData

    def zGetWaveTuple(self):
        """Gets data on all defined wavelengths from the ZEMAX DDE server. This
        function is similar to "zGetWaveDataMatrix()" in MZDDE toolbox.

        zDetWaveTuple() -> waveDataTuple

        ret:
            waveDataTuple: wave data tuple is a 2D tuple with the first
            dimension (first subtuple) containing the wavelengths and the second
            dimension containing the weights like so:
                ((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))

        See also zSetWaveTuple(), zGetWave(), zSetWave()
        """
        waveCount = self.zGetWave(0)[1]
        waveDataTuple = [[],[]]
        for i in range(waveCount):
            cmd = ('GetWave,%i')%(i+1)
            reply = self.conversation.Request(cmd)
            rs = reply.split(',')
            waveDataTuple[0].append(float(rs[0])) # store the wavelength
            waveDataTuple[1].append(float(rs[1])) # store the weight
        return (tuple(waveDataTuple[0]),tuple(waveDataTuple[1]))

    def zInsertConfig(self,number):
        """Insert a new configuration in the multi-configuration editor. The new
        configuration will be placed at the location indicated by the parameter
        `config`.

        zInsertConfig(number)->configNumber

        args:
            number       : (integer) the number of the configuration to insert.
        ret:
            configNumber : (integer) the number of the configuration that is inserted.

        Note:
            1. The configNumber returned could be different from the number in the
               input. This happens, for example, if the "number" is > 1 + number of
               configurations currently in the MCE.
            2. Use zInsertMCO to insert a new multi-configuration operand in the
               multi-configuration editor.

        See also zDeleteConfig.
        """
        return int(self.conversation.Request("InsertConfig,"+str(number)))

    def zInsertMCO(self,operandNumber):
        """Insert a new multi-configuration operand in the multi-configuration editor.

        zInsertMCO(operandNumber)-> retValue

        args:
            operandNumber : (integer) between 1 and the current number of operands
                            plus 1, inclusive.
        ret:
            retValue      : new number of operands.

        See also zDeleteMCO
        """
        return int(self.conversation.Request("InsertMCO,"+str(operandNumber)))

    def zInsertSurface(self,surfNum):
        """Insert a lens surface in the ZEMAX DDE server. The new surface will be
        placed at the location indicated by the parameter surfNum.

        zInsertSruface(surfNum)-> retVal (0 = success)

        See also zSetSurfaceData() to define data for the new surface and the
        zDeleteSurface() functions.
        """
        return int(self.conversation.Request("InsertSurface,"+str(surfNum)))

    def zLoadFile(self,filename,append=None):
        """Loads a ZEMAX file into the server.

        zLoadFile(filename[,append]) -> retVal

        args:
            filename: full path of the ZEMAX file to be loaded. For example:
                      "C:\ZEMAX\Samples\cooke.zmx"
            append (optional): If a non-zero value of append is passed, then the new
                    file is appended to the current file starting at the surface
                    number defined by the value appended.
        retVal:
                0: file successfully loaded
             -999: file could not be loaded (check if the file really exists, or
                   check the path.
             -998: the command timed out
            other: the upload failed.


        See also zSaveFile, zGetPath, zPushLens, zuiLoadFile
        """
        reply = None
        if append==None:
            reply = self.conversation.Request('LoadFile,'+filename)
        else:
            reply = self.conversation.Request('LoadFile,'+filename,append)

        if reply:
            return int(reply) #Note: Zemax returns -999 if update fails.
        else:
            return -998

    def zNewLens(self):
        """Erases the current lens. The "minimum" lens that remains is identical
        to the lens Data Editor when "File,New" is selected. No prompt to save
        the existing lens is given.

        zNewLens-> retVal (retVal = 0 means successful)

        """
        return int(self.conversation.Request('NewLens'))

    def zPushLens(self,timeout = None, updateFlag=None):
        """Copy lens in the ZEMAX DDE server into the Lens Data Editor (LDE).

        zPushLens([timeout,updateFlag]) -> retVal

        args:
            timeout (optional)   : if a timeout in seconds in passed, the client will
                                   wait till the timeout before returning a timeout
                                   error. If no timeout is passed, the default timeout
                                   of 3 seconds is used.
            updateFlag (optional): if 0 or omitted, the open windows are not updated.
                                   if 1, then all open analysis windows are updated.

        retVal:
                0: lens successfully pushed into the LDE.
             -999: the lens could not be pushed into the LDE. (check PushLensPermission)
             -998: the command timed out
            other: the update failed.

        Note that this operation requires the permission of the user running the
        ZEMAX program. The proper use of zPushLens is to first call zPushLensPermission.

        See also zPushLensPermission, zLoadFile, zGetUpdate, zGetPath, zGetRefresh,
        zSaveFile.
        """
        reply = None
        if timeout:
            warnings.warn("Timeout not implemented. Default = 1 min.")
            pass
        if updateFlag==1:
            reply = self.conversation.Request('PushLens,1')
        elif updateFlag == 0 or updateFlag == None:
            reply = self.conversation.Request('PushLens')
        else:
            raise ValueError('Invalid value for flag')

        if reply:
            return int(reply)   #Note, Zemax itself returns -999 if the push lens failed.
        else:
            return -998   #if timeout reached

    def zPushLensPermission(self):
        """Establish if ZEMAX extensions are allowed to push lenses in the LDE.

        zPushLensPermission() -> status

        status:
            1: ZEMAX is set to accept PushLens commands
            0: Extensions are not allowed to use PushLens

        For more details, please refer to the ZEMAX manual.

        See also zPushLens, zGetRefresh
        """
        status = None
        status = self.conversation.Request('PushLensPermission')
        return int(status)

    def zQuickFocus(self,mode,centroid):
        """Performs a quick best focus adjustment for the optical system. The
        "best" focus is chosen as a wavelength weighted average over all fields.

        zQuickFocus(mode,centroid) -> retVal

        arg:
            mode:
                0: RMS spot size
                1: spot x
                2: spot y
                3: wavefront OPD
            centroid: to specify RMS reference
                0: RMS referenced to the chief ray
                1: RMS referenced to image centroid

        ret:
            retVal: 0 for success.
        """
        cmd = ('QuickFocus,%i,%i'%(mode,centroid))
        reply = self.conversation.Request(cmd)
        return reply



    def zSetAperture(self,surfNum,aType,aMin,aMax,xDecenter=0,yDecenter=0,
                                                            apertureFile =' '):
        """Sets aperture details at a ZEMAX lens surface (surface data dialog box).

        zSetAperture(surfNum,aType,aMin,aMax,[xDecenter,yDecenter,apertureFile])
                                           -> apertureInfo

        args:
            surfNum      : surface number (integer)
            aType        : integer code to specify aperture type
                            0 = no aperture (na)
                            1 = circular aperture (ca)
                            2 = circular obscuration (co)
                            3 = spider (s)
                            4 = rectangular aperture (ra)
                            5 = rectangular obscuration (ro)
                            6 = elliptical aperture (ea)
                            7 = elliptical obscuration (eo)
                            8 = user defined aperture (uda)
                            9 = user defined obscuration (udo)
                            10 = floating aperture (fa)
            aMin         : min radius(ca), min radius(co),width of arm(s),
                           X-half width(ra),X-half width(ro),X-half width(ea),
                           X-half width(eo)
            aMax         : max radius(ca), max radius(co),number of arm(s),
                           X-half width(ra),X-half width(ro),X-half width(ea),
                           X-half width(eo)
                           see "Aperture type and other aperture controls" for more details.
            xDecenter    : amount of decenter from current optical axis (lens units)
            yDecenter    : amount of decenter from current optical axis (lens units)
            apertureFile : a text file with .UDA extention. see "User defined
                           apertures and obscurations" in ZEMAX manual for more details.

        ret:
            apertureInfo: apertureInfo is a tuple containing the following:
                aType     : (see above)
                aMin      : (see above)
                aMax      : (see above)
                xDecenter : (see above)
                yDecenter : (see above)

        Example:
        apertureInfo = zSetAperture(2,1,5,10,0.5,0,'apertureFile.uda')
        or
        apertureInfo = zSetAperture(2,1,5,10)

        See also zGetAperture()
        """
        cmd = ('SetAperture,%i,%i,%1.20,%1.20g,%1.20g,%1.20g,%s'
               %(surfNum,aType,aMin,aMax,xDecenter,yDecenter,apertureFile))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        apertureInfo = tuple([float(elem) for elem in rs])
        return apertureInfo

    def zSetExtra(self,surfaceNumber,columnNumber,value):
        """Sets extra surface data (value) in the Extra Data Editor for the surface
        indicatd by surfaceNumber.

        zSetExtra(surfaceNumber,columnNumber,value)->retValue

        arg:
            surfaceNumber : (integer) surface number
            columnNumber  : (integer) column number
            value         : (float) value
        ret:
            retValue      : (float) numeric data value

        See also zGetExtra
        """
        cmd = ('SetExtra,%i,%i,%1.20g')%(surfaceNumber,columnNumber,value)
        reply = self.conversation.Request(cmd)
        return float(reply)

    def zSetField(self,n,arg1,arg2,arg3=1.0,vdx=0.0,vdy=0.0,vcx=0.0,vcy=0.0,van=0.0):
        """Sets the field data for a particular field point.

        There are 2 ways of using this function:

            zSetField(0, fieldType,totalNumFields,fieldNormalization)-> fieldData
             OR
            zSetField(n,xf,yf [,wgt,vdx,vdy,vcx,vcy,van])-> fieldData

        args[if n == 0]:
            0         : to set general field parameters
            arg1 : the field type
                  0 = angle, 1 = object height, 2 = paraxial image height, and
                  3 = real image height
            arg2 : total number of fields
            arg3 : normalization type [0=radial, 1=rectangular(default)]

        args[if 0 < n <= number of fields]:
            arg1 (fx),arg2 (fy) : the field x and field y values
            arg3 (wgt)          : field weight (default = 1.0)
            vdx,vdy,vcx,vcy,van : vignetting factors (default = 0.0), See below.

        ret [if n=0]: fieldData is a tuple containing the following
            type                 : integer (0=angles in degrees, 1=object height
                                            2=paraxial image height,
                                            3=real image height)
            number               : number of fields currently defined
            max_x_field          : values used to normalize x field coordinate
            max_y_field          : values used to normalize y field coordinate
            normalization_method : field normalization method (0=radial, 1=rectangular)

        ret [if 0 < n <= number of fields]: fieldData is a tuple containing the following
            xf     : the field x value
            yf     : the field y value
            wgt    : field weight
            vdx    : decenter x vignetting factor
            vdy    : decenter y vignetting factor
            vcx    : compression x vignetting factor
            vcy    : compression y vignetting factor
            van    : angle vignetting factor

        Note: the returned tuple's content and structure is exactly same as that
        of zGetField()

        See also zGetField()
        """
        if n:
            cmd = ('SetField,%i,%1.20g,%1.20g,%1.20g,%1.20g,%1.20g,%1.20g,%1.20g,%1.20g'
                         %(n,arg1,arg2,arg3,vdx,vdy,vcx,vcy,van))
        else:
            cmd = ('SetField,%i,%i,%i,%i'%(0,arg1,arg2,arg3))

        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        if n:
            fieldData = tuple([float(elem) for elem in rs])
        else:
            fieldData = tuple([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return fieldData

    def zSetFieldTuple(self,fieldType,fNormalization, iFieldDataTuple):
        """Sets all field points from a 2D field tuple structure. This function
        is similar to the function "zSetFieldMatrix()" in MZDDE toolbox.

        zSetFieldTuple(fieldType,fNormalization, iFieldDataTuple)->oFieldDataTuple

        args:
            fieldType: the field type (0=angle, 1=object height, 2=paraxial image
                       height, and 3 = real image height
            fNormalization: field normalization (0=radial, 1=rectangular)
            iFieldDataTuple = the input field data tuple is an N-D tuple (0<N<=12)
            with every dimension representing a single field location. It can be
            constructed as shown here with an example:

            iFieldDataTuple =
            ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # xf=0.0,yf=0.0,wgt=1.0,vdx=vdy=vcx=vcy=van=0.0
             (0.0,5.0,1.0),                     # xf=0.0,yf=5.0,wgt=1.0
             (0.0,10.0))                        # xf=0.0,yf=10.0
        ret:
            oFieldDataTuple: the output field data tuple is also a N-D tuple similar
            to the iFieldDataTuple, except that for each field location all 8
            field parameters are returned.

        See also zSetField(), zGetField(), zGetFieldTuple()
        """
        fieldCount = len(iFieldDataTuple)
        if not 0 < fieldCount <= 12:
            raise ValueError('Invalid number of fields')
        cmd = ('SetField,%i,%i,%i,%i'%(0,fieldType,fieldCount,fNormalization))
        reply = self.conversation.Request(cmd)
        oFieldDataTuple = [ ]
        for i in range(fieldCount):
            fieldData = self.zSetField(i+1,*iFieldDataTuple[i])
            oFieldDataTuple.append(fieldData)
        return tuple(oFieldDataTuple)

    def zSetPrimaryWave(self,primaryWaveNumber):
        """Sets the wavelength data in the ZEMAX DDE server. This function emulates
        the function "zSetPrimaryWave()" of the MZDDE toolbox.

        zSetPrimaryWave(primaryWaveNumber) -> waveData

        args:
            primaryWaveNumber: the wave number to set as primary

          ret: waveData is a tuple containing the following:
            primary : number indicating the primary wavelength (integer)
            number  : number of wavelengths currently defined (integer).

        Note: the returned tuple is exactly same in structure and contents to that
        returned by zGetWave(0).

        See also zSetWave(), zSetWave(), zSetWaveTuple(), zGetWaveTuple().
        """
        waveData = self.zGetWave(0)
        cmd = ('SetWave,%i,%i,%i')%(0,primaryWaveNumber,waveData[1])
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        waveData = tuple([int(elem) for elem in rs])
        return waveData

    def zSetSurfaceData(self,surfaceNumber,code,value,arg2=None):
        """Sets surface data on a sequential lens surface.

        zSetSurfaceData(surfaceNum,code,value [, arg2])-> surfaceDatum

        args:
            surfaceNum : the surface number
            code       : integer number (see below)
            value      : string type if code = 0,1,4,7 or 9  else float type
            arg2       : (Optional) for item codes above 70.



        Sets surface datum at surfaceNumber depending on the code according to
        the following table.
        The "value" is the required value to which the datum should be set.
        Supply a string or a numeric Value according to the following table.
        arg2 is required for item codes above 70.

        To set the surface type to a user defined surface, send the new DLL name
        using code 9 rather by setting the surface type.

        Code      - Datum to be set by by zSetSurfaceData
        0         - Surface type name. (string)
        1         - Comment. (string)
        2         - Curvature (numeric).
        3         - Thickness. (numeric)
        4         - Glass. (string)
        5         - Semi-Diameter. (numeric)
        6         - Conic. (numeric)
        7         - Coating. (string)
        8         - Thermal Coefficient of Expansion (TCE).
        9         - User-defined .dll (string)
        20        - Ignore surface flag, 0 for not ignored, 1 for ignored.
        51        - Tilt, Decenter order before surface; 0 for Decenter then
                    Tilt, 1 for Tilt then Decenter.
        52        - Decenter x
        53        - Decenter y
        54        - Tilt x before surface
        55        - Tilt y before surface
        56        - Tilt z before surface
        60        - Status of Tilt/Decenter after surface. 0 for explicit, 1
                    for pickup current surface,2 for reverse current surface,
                    3 for pickup previous surface, 4 for reverse previous
                    surface,etc.
        61        - Tilt, Decenter order after surface; 0 for Decenter then
                    Tile, 1 for Tilt then Decenter.
        62        - Decenter x after surface
        63        - Decenter y after surface
        64        - Tilt x after surface
        65        - Tilt y after surface
        66        - Tilt z after surface
        70        - Use Layer Multipliers and Index Offsets. Use 1 for true,
                    0 for false.
        71        - Layer Multiplier value. The coating layer number is defined
                    by arg2.
        72        - Layer Multiplier status. Use 0 for fixed, 1 for variable,
                    or n+1 for pickup from layer n. The coating layer number
                    is defined by arg2.
        73        - Layer Index Offset value. The coating layer number is
                    defined by arg2.
        74        - Layer Index Offset status. Use 0 for fixed, 1 for variable,
                    or n+1 for pickup from layer n.The coating layer number is
                    defined by arg2.
        75        - Layer Extinction Offset value. The coating layer number is
                    defined by arg2.
        76        - Layer Extinction Offset status. Use 0 for fixed, 1 for
                    variable, or n+1 for pickup from layer n. The coating layer
                    number is defined by arg2.
        Other     - Reserved for future expansion of this feature.

        See also zGetSurfaceData and ZemaxSurfTypes
        """
        cmd = ('SetSurfaceData,%i,%i')%(surfaceNumber,code)
        if code in (0,1,4,7,9):
            if isinstance(value,str):
                cmd = cmd+','+value
            else:
                raise ValueError('Invalid input, expecting string type code')
        else:
            if not isinstance(value,str):
                cmd = cmd+','+str(value)
            else:
                raise ValueError('Invalid input, expecting float type code')
        if code > 70:
            if arg2 != None:
                cmd = cmd+","+str(arg2)
            else:
                raise ValueError('Invalid input, expecting argument')
        reply = self.conversation.Request(cmd)
        if code in (0,1,4,7,9):
            surfaceDatum = reply.split()[0]
        else:
            surfaceDatum = float(reply)
        return surfaceDatum

    def zSetSurfaceParameter(self,surfaceNumber,parameter,value):
        """Set surface parameter data.
        zSetSurfaceParameter(surfaceNumber, parameter, value)-> parameterData

        args:
            surfaceNumber  : (integer) surface number of the surface
            parameter      : (integer) parameter (Par in LDE) number being set
        ret:
            parameterData  : (float) the parameter value

        See also zSetSurfaceData, zGetSurfaceParameter
        """
        cmd = ('SetSurfaceParameter,%i,%i,%1.20g')%(surfaceNumber,parameter,value)
        reply = self.conversation.Request(cmd)
        return float(reply)


    def zSetSystem(self,unitCode,stopSurf,rayAimingType,useEnvData,
                                              temp,pressure,globalRefSurf):
        """Sets a number of general systems property (General Lens Data)

        zSetSystem(unitCode,stopSurf,rayAimingType,useenvdata,
                     temp,pressure,globalRefSurf) -> systemData

        args:
            unitCode      : lens units code (0,1,2,or 3 for mm, cm, in, or M)
            stopSurf      : the stop surface number
            rayAimingType : ray aiming type (0,1, or 2 for off, paraxial or real)
            useEnvData    : use environment data flag (0 or 1 for no or yes) [ignored]
            temp          : the current temperature
            pressure      : the current pressure
            globalRefSurf : the global coordinate reference surface number

        ret:
            systemData : the systemData is a tuple with the following elements:
                numSurfs      : number of surfaces
                unitCode      : lens units code (0,1,2,or 3 for mm, cm, in, or M)
                stopSurf      : the stop surface number
                nonAxialFlag  : flag to indicate if system is non-axial symmetric
                                (0 for axial, 1 if not axial)
                rayAimingType : ray aiming type (0,1, or 2 for off, paraxial or real)
                adjust_index  : adjust index data to environment flag (0 if false, 1 if true)
                temp          : the current temperature
                pressure      : the current pressure
                globalRefSurf : the global coordinate reference surface number
                need_save     : indicates whether the file has been modified. [Deprecated]

        Note: the returned data structure is exactly similar to the data structure
        returned by the zGetSystem() method.

        If you are interested in setting the system apeture, such as aperture type,
        aperture value, etc, use zSetSystemAper().

        See also zGetSystem, zGetSystemAper, zSetSystemAper, zGetAperture, zSetAperture
        """
        cmd = ('SetSystem,%i,%i,%i,%i,%1.20g,%1.20g,%i'
               %(unitCode,stopSurf,rayAimingType,useEnvData,temp,pressure,
               globalRefSurf))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        systemData = tuple([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zSetSystemAper(self,aType,stopSurf,apertureValue):
        """Sets the lens system aperture and corresponding data.

        zSetSystemAper(aType, stopSurf, apertureValue)-> systemAperData

        arg:
            aType              : integer indicating the system aperture
                                 0 = entrance pupil diameter
                                 1 = image space F/#
                                 2 = object space NA
                                 3 = float by stop
                                 4 = paraxial working F/#
                                 5 = object cone angle
            stopSurf           : stop surface
            value              : if aperture type == float by stop
                                     value is stop surface semi-diameter
                                 else
                                     value is the sytem aperture

        ret:
            systemAperData: systemAperData is a tuple containing the following
                aType              : (see above)
                stopSurf           : (see above)
                value              : (see above)

        Note: the returned tuple is the same as the returned tuple of zGetSystemAper()

        See also, zGetSystem(), zGetSystemAper()
        """
        cmd = ('SetSystemAper,%i,%i,%1.20g' %(aType,stopSurf,apertureValue))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        systemAperData = tuple([float(elem) for elem in rs])
        return systemAperData

    def zSetWave(self,n,arg1,arg2):
        """Sets the wavelength data in the ZEMAX DDE server.

        There are 2 ways to use this function:
            zSetWave(0,primary,number) -> waveData
             OR
            zSetWave(n,wavelength,weight) -> waveData

            args [if n==0]:
                0             : if n=0, the function sets general wavelength data
                primary (arg1): primary wavelength value to set
                number (arg2) : total number of wavelengths to set

            args [if 0 < n <= number of wavelengths]:
                n                : wavelength number to set
                wavelength (arg1): wavelength in micrometers (floating)
                weight    (arg2) : weight (floating)

            ret:
                if n==0: waveData is a tuple containing the following:
                    primary : number indicating the primary wavelength (integer)
                    number  : number of wavelengths currently defined (integer).
                elif 0 < n <= number of wavelengths: waveData consists of:
                    wavelength : value of the specific wavelength (floating point)
                    weight     : weight of the specific wavelength (floating point)

        Note: the returned tuple is exactly same in structure and contents to that
        returned by zGetWave().

        See also zGetWave(), zSetPrimaryWave(), zSetWaveTuple(), zGetWaveTuple().
        """
        if n:
            cmd = ('SetWave,%i,%1.20g,%1.20g')%(n,arg1,arg2)
        else:
            cmd = ('SetWave,%i,%i,%i')%(0,arg1,arg2)
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        if n:
            waveData = tuple([float(ele) for ele in rs])
        else:
            waveData = tuple([int(ele) for ele in rs])
        return waveData

    def zSetWaveTuple(self,iWaveDataTuple):
        """Sets wavelength and weight data from a matrix. This function is similar
        to the function "zSetWaveMatrix()" in the MZDDE toolbox.

        zSetWaveTuple(iWaveDataTuple)-> oWaveDataTuple

        arg:
            iWaveDataTuple: the input wave data tuple is a 2D tuple with the first
            dimension (first subtuple) containing the wavelengths and the second
            dimension containing the weights like so:
                ((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))
            The first wavelength (wave1) is assigned to be the primary wavelength.
            To change the primary wavelength use zSetWavePrimary()

        ret:
            oWaveDataTuple: the output wave data tuple is also a 2D tuple similar
            to the iWaveDataTuple.

        See also zGetWaveTuple(), zSetWave(), zSetWavePrimary()
        """
        waveCount = len(iWaveDataTuple[0])
        oWaveDataTuple = [[],[]]
        self.zSetWave(0,1,waveCount) # Set no. of wavelen & the wavelen to 1
        for i in range(waveCount):
            cmd = ('SetWave,%i,%1.20g,%1.20g')%(i+1,iWaveDataTuple[0][i],
                                                        iWaveDataTuple[1][i])
            reply = self.conversation.Request(cmd)
            rs = reply.split(',')
            oWaveDataTuple[0].append(float(rs[0])) # store the wavelength
            oWaveDataTuple[1].append(float(rs[1])) # store the weight
        return (tuple(oWaveDataTuple[0]),tuple(oWaveDataTuple[1]))


# ****************************************************************
#                      CONVENIENCE FUNCTIONS
# ****************************************************************
    def spiralSpot(self,hy,hx,waveNum,spirals,rays,mode=0):
        """Convenience function to produce a series of x,y values of rays traced
        in a spiral over the entrance pupil to the image surface. i.e. the final
        destination of the rays is the image surface. This function imitates its
        namesake from MZDDE toolbox.

        spiralSpot(hy,hx,waveNum,spirals,rays[,mode])->[x,y,z,intensity]

        Note: Since the spiralSpot function performs a GetRefresh() to load lens
        data from the LDE to the DDE server, perform PushLens() before calling
        spiralSpot.
        """
        status = self.zGetRefresh()
        if ~status:
            finishAngle = spirals*2*pi
            dTheta = finishAngle/(rays-1)
            theta = [i*dTheta for i in range(rays)]
            r = [i/finishAngle for i in theta]
            px = [r[i]*cos(theta[i]) for i in range(len(theta))]
            py = [r[i]*sin(theta[i]) for i in range(len(theta))]
            x = [] # x-coordinate of the image surface
            y = [] # y-coordinate of the image surface
            z = [] # z-coordinate of the image surface
            intensity = [] # the relative transmitted intensity of the ray
            for i in range(len(px)):
                rayTraceData = self.zGetTrace(waveNum,mode,-1,hx,hy,px[i],py[i])
                if rayTraceData[0] == 0:
                    x.append(rayTraceData[2])
                    y.append(rayTraceData[3])
                    z.append(rayTraceData[4])
                    intensity.append(rayTraceData[11])
                else:
                    print "Raytrace Error"
                    exit()
                    #FIXME raise an error here
            return [x,y,z,intensity]
        else:
            print "Couldn't copy lens data from LDE to server, no tracing can be performed"
            return [None,None]

    def lensScale(self,factor=2.0,ignoreSurfaces=None):
        """Scale the lens design by factor specified.

        lensScale([factor,ignoreSurfaces])->ret

        args:
            factor         : the scale factor. If no factor are passed, the design
                             will be scaled by a factor of 2.0
            ignoreSurfaces : (tuple) of surfaces that are not to be scaled. Such as
                             (0,2,3) to ignore surfaces 0 (object surface), 2 and
                             3. Or (OBJ,2, STO,IMG) to ignore object surface, surface
                             number 2, stop surface and image surface.
        ret:
            0 : success
            1 : success with warning
            -1: failure

        Notes:
            1. WARNING: this function implementation is not yet complete.
                * Note all surfaces have been implemented
                * ignoreSurface option has not been implemented yet.

        Limitations:
            1. Cannot scale pupil shift x,y, and z in the General settings as Zemax
               hasn't provided any command to do so using the DDE. The pupil shift
               values are also scaled, when a lens design is scaled, when the ray-
               aiming is on. However, this is not a serious limitation for most cases.
        """
        ret = 0 # assuming successful return
        lensFile = self.zGetFile()
        if factor == 1:
            return ret
        #Scale the "system aperture" appropriately
        sysAperData = self.zGetSystemAper()
        if sysAperData[0] == 0:   # System aperture if EPD
            stopSurf = sysAperData[1]
            aptVal = sysAperData[2]
            self.zSetSystemAper(0,stopSurf,factor*aptVal)
        elif sysAperData[0] in (1,2,4): # Image Space F/#, Object Space NA, Working Para F/#
            ##print "Warning: Scaling of aperture may be incorrect"
            pass
        elif sysAperData[0] == 3: # System aperture if float by stop
            pass
        elif sysAperData[0] == 5: # Object Cone Angle
            print "Warning: Scaling for Object Cone Angle aperture type may be incorrect for %s" %(lensFile)
            ret = 1
        #Get the number of surfaces
        numSurf = 0
        recSystemData_g = self.zGetSystem() #Get the current system parameters
        numSurf = recSystemData_g[0]
        #print "Number of surfaces in the lens: ", numSurf

        if recSystemData_g[4] > 0:
            print "Warning: Ray aiming is ON in %s. But cannot scale Pupil Shift values." %(lensFile)

        #Scale individual surface properties in the LDE
        for surfNum in range(0,numSurf+1): #Start from the object surface ... to scale thickness if not infinity
            #Scale the basic data common to all surface types such as radius, thickness
            #and semi-diameter
            surfName = self.zGetSurfaceData(surfNum,0)
            curv = self.zGetSurfaceData(surfNum,2)
            thickness = self.zGetSurfaceData(surfNum,3)
            semiDiam = self.zGetSurfaceData(surfNum,5)
            ##print "Surf#:",surfNum,"Name:",surfName,"Curvature:",curv,"Thickness:",thickness,"Semi-Diameter:", semiDiam
            #scale the basic data
            scaledCurv = self.zSetSurfaceData(surfNum,2,curv/factor)
            if thickness < 1.0E+10: #Scale the thickness if it not Infinity (-1.0E+10 in Zemax)
                scaledThickness = self.zSetSurfaceData(surfNum,3,factor*thickness)
            scaledSemiDiam = self.zSetSurfaceData(surfNum,5,factor*semiDiam)
            ##print "scaled", surfNum,surfName,scaledCurv,scaledThickness,scaledSemiDiam

            #scaling parameters of surface individually
            if surfName == 'STANDARD': #Std surface - plane, spherical, or conic aspheric
                pass #Nothing to do
            elif surfName in ('BINARY_1','BINARY_2'):
                binSurMaxNum = {'BINARY_1':233,'BINARY_2':243}
                for pNum in range(1,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    par_ret = self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*pNum)*par)
                #Scale norm radius in the extra data editor
                epar2 = self.zGetExtra(surfNum,2) #Norm radius
                epar2_ret = self.zSetExtra(surfNum,2,factor*epar2)
                #scale the coefficients of the Zernike Fringe polynomial terms in the EDE
                numBTerms = int(self.zGetExtra(surfNum,1))
                if numBTerms > 0:
                    for i in range(3,binSurMaxNum[surfName]): #scaling of terms 3 to 232, p^480 for Binary1 and Binary 2 respectively
                        if i > numBTerms + 2: #(+2 because the terms starts from par 3)
                            break
                        else:
                            epar = self.zGetExtra(surfNum,i)
                            epar_ret = self.zSetExtra(surfNum,i,factor*epar)
            elif surfName == 'BINARY_3':
                #Scaling of parameters in the LDE
                par1 = self.zGetSurfaceParameter(surfNum,1) # R2
                par1_ret = self.zSetSurfaceParameter(surfNum,1,factor*par1)
                par4 = self.zGetSurfaceParameter(surfNum,4) # A2, need to scale A2 before A1, because A2>A1>0.0 always
                par4_ret = self.zSetSurfaceParameter(surfNum,4,factor*par4)
                par3 = self.zGetSurfaceParameter(surfNum,3) # A1
                par3_ret = self.zSetSurfaceParameter(surfNum,3,factor*par3)
                numBTerms = int(self.zGetExtra(surfNum,1))    #Max possible is 60
                for i in range(2,243,4):  #242
                    if i > 4*numBTerms + 1: #(+1 because the terms starts from par 2)
                        break
                    else:
                        par_r1 = self.zGetExtra(surfNum,i)
                        par_r1_ret = self.zSetExtra(surfNum,i,par_r1/factor**(i/2))
                        par_p1 = self.zGetExtra(surfNum,i+1)
                        par_p1_ret = self.zSetExtra(surfNum,i+1,factor*par_p1)
                        par_r2 = self.zGetExtra(surfNum,i+2)
                        par_r1_ret = self.zSetExtra(surfNum,i+2,par_r2/factor**(i/2))
                        par_p2 = self.zGetExtra(surfNum,i+3)
                        par_p2_ret = self.zSetExtra(surfNum,i+3,factor*par_p2)

            elif surfName == 'COORDBRK': #Coordinate break,
                par = self.zGetSurfaceParameter(surfNum,1) # decenter X
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
                par = self.zGetSurfaceParameter(surfNum,2) # decenter Y
                par_ret = self.zSetSurfaceParameter(surfNum,2,factor*par)
            elif surfName == 'EVENASPH': #Even Asphere,
                for pNum in range(1,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    par_ret = self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*pNum)*par)
            elif surfName == 'GRINSUR1': #Gradient1
                par1 = self.zGetSurfaceParameter(surfNum,1) #Delta T
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par1)
                par3 = self.zGetSurfaceParameter(surfNum,3) #coeff of radial quadratic index
                par_ret = self.zSetSurfaceParameter(surfNum,3,par3/(factor**2))
                par4 = self.zGetSurfaceParameter(surfNum,4) #index of radial linear index
                par_ret = self.zSetSurfaceParameter(surfNum,4,par4/factor)
            elif surfName == 'GRINSUR9': #Gradient9
                par = self.zGetSurfaceParameter(surfNum,1) #Delta T
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'GRINSU11': #Grid Gradient surface with 1 parameter
                par = self.zGetSurfaceParameter(surfNum,1) #Delta T
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'PARAXIAL': #Paraxial
                par = self.zGetSurfaceParameter(surfNum,1) #Focal length
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'PARAX_XY': #Paraxial XY
                par = self.zGetSurfaceParameter(surfNum,1) # X power
                par_ret = self.zSetSurfaceParameter(surfNum,1,par/factor)
                par = self.zGetSurfaceParameter(surfNum,2) # Y power
                par_ret = self.zSetSurfaceParameter(surfNum,2,par/factor)
            elif surfName == 'PERIODIC':
                par = self.zGetSurfaceParameter(surfNum,1) #Amplitude/ peak to valley height
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
                par = self.zGetSurfaceParameter(surfNum,2) #spatial frequency of oscillation in x
                par_ret = self.zSetSurfaceParameter(surfNum,2,par/factor)
                par = self.zGetSurfaceParameter(surfNum,3) #spatial frequency of oscillation in y
                par_ret = self.zSetSurfaceParameter(surfNum,3,par/factor)
            elif surfName == 'POLYNOMI':
                for pNum in range(1,5): # from Par 1 to Par 4 for x then Par 5 to Par 8 for y
                    parx = self.zGetSurfaceParameter(surfNum,pNum)
                    pary = self.zGetSurfaceParameter(surfNum,pNum+4)
                    parx_ret = self.zSetSurfaceParameter(surfNum,pNum,
                                                      factor**(1-2.0*pNum)*parx)
                    pary_ret = self.zSetSurfaceParameter(surfNum,pNum+4,
                                                         factor**(1-2.0*pNum)*pary)
            elif surfName == 'TILTSURF': #Tilted surface
                pass           #No parameters to scale
            elif surfName == 'TOROIDAL':
                par = self.zGetSurfaceParameter(surfNum,1) #Radius of rotation
                par_ret = self.zSetSurfaceParameter(surfNum,1,factor*par)
                for pNum in range(2,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    par_ret = self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*(pNum-1))*par)
                #scale parameters from the extra data editor
                epar = self.zGetExtra(surfNum,2)
                epar_ret = self.zSetExtra(surfNum,2,factor*epar)
            elif surfName == 'FZERNSAG': # Zernike fringe sag
                for pNum in range(1,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    par_ret = self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*pNum)*par)
                par9      = self.zGetSurfaceParameter(surfNum,9) # decenter X
                par9_ret  = self.zSetSurfaceParameter(surfNum,9,factor*par9)
                par10     = self.zGetSurfaceParameter(surfNum,10) # decenter Y
                par10_ret = self.zSetSurfaceParameter(surfNum,10,factor*par10)
                #Scale norm radius in the extra data editor
                epar2 = self.zGetExtra(surfNum,2) #Norm radius
                epar2_ret = self.zSetExtra(surfNum,2,factor*epar2)
                #scale the coefficients of the Zernike Fringe polynomial terms in the EDE
                numZerTerms = int(self.zGetExtra(surfNum,1))
                if numZerTerms > 0:
                    epar3 = self.zGetExtra(surfNum,3) #Zernike Term 1
                    epar3_ret = self.zSetExtra(surfNum,3,factor*epar3)
                    #Zernike terms 2,3,4,5 and 6 are not scaled.
                    for i in range(9,40): #scaling of Zernike terms 7 to 37
                        if i > numZerTerms + 2: #(+2 because the Zernike terms starts from par 3)
                            break
                        else:
                            epar = self.zGetExtra(surfNum,i)
                            epar_ret = self.zSetExtra(surfNum,i,factor*epar)
            else:
                print "Warning: Scaling for surface type %s in file %s not implemented!!" %(surfName,lensFile)
                ret = -1
                pass

        #Scale appropriate parameters in the Multi-configuration editor, such as THIC, APER ...
        #maybe, use GetConfig(), SetConfig() and GetMulticon

        #Scale appropriate parameters in the Tolerance Data Editor

        #Scale the parameters in the Field data Editor if the field positions are
        #NOT of angle type.
        (fType,fNum,fxMax,fyMax,fNorm) = self.zGetField(0)
        if fType != 0:
            fieldDataTuple = self.zGetFieldTuple()
            fieldDataTupleScaled = []
            for i in range(fNum):
                tField = list(fieldDataTuple[i])   #have to convert into a list as a tuple is immutable.
                tField[0],tField[1] = factor*tField[0],factor*tField[1]
                fieldDataTupleScaled.append(tuple(tField))
            fieldDataTupleScaled = self.zSetFieldTuple(fType,fNorm,
                                                 tuple(fieldDataTupleScaled))
        return ret


    def calculateHiatus(self,txtFileName2Use=None,keepFile=False):
        """Function to calculate the Hiatus, also known as the Null space, or nodal
           space, or the interstitium (i.e. the distance between the two principal
           planes.

        calculateHiatus([txtFileName2Use,keepFile])-> hiatus

        args:
            txtFileName2Use : (optional, string) If passed, the prescription file
                              will be named such. Pass a specific txtFileName if
                              you want to dump the file into a separate directory.
            keepFile        : (optional, bool) If false (default), the prescription
                              file will be deleted after use. If true, the file
                              will persist.
        ret:
            hiatus          : the value of the hiatus

        Note:
            1. Decision to use Prescription file:
               The cardinal points information is retrieved from the prescription
               file. One could also request Zemax to write just the "Cardinal Points"
               data to a text file. In fact, such a file is much smaller and it
               also provides information about the number of surfaces. In most
               situations, especially if only cardinal points'/planes' information
               is required, one may just use "zGetTextFile(textFileName,'Car',"None",0)".
               Zemax then calculates the cardianl points/planes only for the primary
               wavelength. However, the file obtained in the latter method doesn't
               retrieve information about the distances of the first surface and
               the image surface required for calculating the hiatus.
            2. Decision to read all lines from the files into a list:
               It is very difficult (if not impossible) to read the prescirption
               files using bytes as we want to get to a specific position based
               on "keywords" and not "bytes". (we are not guaranteed to find the
               same "keyword" for a specific byte-based-position everytime we read
               a prescription file). If we read the file line-by-line such as
               "for line in file" (using the file iterable object) it becomes hard to
               read, identify and store a specific line which doesn't have any identifiable
               keywords. Also, because of possible data loss, Python raises an exception,
               if we try to use readline() or readlines() within the "for line in file"
               iteration.
        """
        if txtFileName2Use != None:
            textFileName = txtFileName2Use
        else:
            cd = os.getcwd()
            textFileName = cd +"\\"+"prescriptionFile.txt"
        ret = self.zGetTextFile(textFileName,'Pre',"None",0)
        assert ret == 0
        recSystemData_g = self.zGetSystem() #Get the current system parameters
        numSurf       = recSystemData_g[0]
        #Open the text file in read mode to read
        fileref = open(textFileName,"r")
        principalPlane_objSpace = 0.0; principalPlane_imgSpace = 0.0; hiatus = 0.0
        count = 0
        #The number of expected Principal planes in each Pre file is equal to the
        #number of wavelengths in the general settings of the lens design
        #See Note 2 for the reasons why the file was not read as an iterable object
        #and instead, we create a list of all the lines in the file, which is obviously
        #very wasteful of memory
        line_list = fileref.readlines()
        fileref.close()

        for line_num,line in enumerate(line_list):
            #Extract the image surface distance from the global ref sur (surface 1)
            sectionString = "GLOBAL VERTEX COORDINATES, ORIENTATIONS, AND ROTATION/OFFSET MATRICES:"
            if line.rstrip()== sectionString:
                ima_3 = line_list[line_num + numSurf*4 + 6]
                ima_z = float(ima_3.split()[3])

            #Extract the Principal plane distances.
            if "Principal Planes" in line and "Anti" not in line:
                principalPlane_objSpace += float(line.split()[3])
                principalPlane_imgSpace += float(line.split()[4])
                count +=1  #Increment (wavelength) counter for averaging

        #Calculate the average (for all wavelengths) of the principal plane distances
        if count > 0:
            principalPlane_objSpace = principalPlane_objSpace/count
            principalPlane_imgSpace = principalPlane_imgSpace/count
            #Calculate the hiatus (only if count > 0) as
            #hiatus = (img_surf_dist + img_surf_2_imgSpacePP_dist) - objSpacePP_dist
            hiatus = abs(ima_z + principalPlane_imgSpace - principalPlane_objSpace)

        if not keepFile:
            #Delete the prescription file (the directory remains clean)
            os.remove(textFileName)
        return hiatus






# ***************************************************************************
#                      TEST THE FUNCTIONS
# ***************************************************************************
# Initial code testing used to be done using the following lines of code.
# Currently all functionality are being tested using the unit test module.
# The following lines are left for quick test. The code will not be invoked if
# the module is imported! In order to execute the test_PyZDDE() function, "run"
# this (pyZDDE.py) file. It may prove to be useful to quickly test your system.

def test_PyZDDE():
    "Test the pyZDDE module functions"
    zmxfp = os.getcwd()+'\\ZMXFILES\\'
    # Create PyZDDE object(s)
    link0 = pyzdde()
    link1 = pyzdde()
    link2 = pyzdde()  # this object shall be deleted randomly

    print "\nTEST: zDDEInit()"
    print "---------------"
    status = link0.zDDEInit()
    print "Status for link 0:", status
    assert status == 0
    print "App Name for Link 0:", link0.appName
    print "Connection status for Link 0:", link0.connection
    print "\n"
    time.sleep(0.25)   # Not required, but just for observation

    #link1 = pyzdde()
    status = link1.zDDEInit()
    print "Status for link 1:",status
    #assert status == 0   # In older versions of Zemax, unable to create second
                          # communication link.
    if status != 0:
        warnings.warn("Couldn't create second channel.\n")
    else:
        print "App Name for Link 1:", link1.appName
        print "Connection status for Link 1:", link1.connection
        print "\n"
    time.sleep(0.25)   # Not required, but just for observation

    print "\nTEST: zGetDate()"
    print "----------------"
    print "Date: ", link0.zGetDate().rstrip()  # strip off the newline char

    print "\nTEST: zGetSerial()"
    print "------------------"
    ser = link0.zGetSerial()
    print "Serial #:", ser

    print "\nTEST: zGetVersion()"
    print "----------------"
    print "version number: ", link0.zGetVersion()

    print "\nTEST: zSetTimeout()"
    print "------------------"
    link0.zSetTimeout(3)

    #Delete link2 randomly
    print "\nTEST: Random deletion of object"
    print "--------------------------------"
    print "Deleting object link2"
    del link2

    print "\nTEST: zLoadFile()"
    print "-------------------"
    filename = zmxfp+"nonExistantFile.zmx"
    ret = link0.zLoadFile(filename)
    assert ret == -999
    filename = zmxfp+"Cooke 40 degree field.zmx"
    ret = link0.zLoadFile(filename)
    assert ret == 0
    print "zLoadFile test successful"

    if link1.connection:
        print "\nTEST: zLoadFile() @ link 1 (second channel)"
        print "-------------------"
        filename = zmxfp+"Double Gauss 5 degree field.zmx"
        assert ret == 0
        print "zLoadFile test @ link 1 successful"

    print "\nTEST: zPushLensPermission()"
    print "---------------------------"
    status = link0.zPushLensPermission()
    if status:
        print "Extensions are allowed to push lens."

        print "\nTEST: zPushLens()"
        print "-----------------"
        # First try to push a lens with invalid flag argument
        try:
            ret = link0.zPushLens(updateFlag=10)
        except:
            info = exc_info()
            print "Exception error:", info[0]
            #assert info[0] == 'exceptions.ValueError'
            assert cmp(str(info[0]),"<type 'exceptions.ValueError'>") == 0

        # TEST ALL FUNCTIONS THAT REQUIRE PUSHLENS() ... HERE!
        #Push lens without any parameters
        ret = link0.zPushLens()
        if ret ==0:
            print "Lens update without any arguments suceeded. ret value = ", ret
        else:
            print "Lens update without any arguments FAILED. ret value = ", ret
        #Push lens with some valid parameters
        ret = link0.zPushLens(updateFlag=1)
        if ret == 0:
            print "Lens update with flag=1 suceeded. ret value = ", ret
        else:
            print "Lens update with flag=1 FAILED. ret value = ", ret

    else: # client do not have permission to push lens
        print "Extensions are not allowed to push lens. Please enable it."

    #Continue with other tests
    print "\nTEST: zGetTrace()"
    print "------------------"
    rayTraceData = link0.zGetTrace(3,0,5,0.0,1.0,0.0,0.0)
    (errorCode,vigCode,x,y,z,l,m,n,l2,m2,n2,intensity) = link0.zGetTrace(3,0,5,
                                                               0.0,1.0,0.0,0.0)
    assert rayTraceData[0]  == errorCode
    assert rayTraceData[1]  == vigCode
    assert rayTraceData[2]  == x
    assert rayTraceData[3]  == y
    assert rayTraceData[4]  == z
    assert rayTraceData[5]  == l
    assert rayTraceData[6]  == m
    assert rayTraceData[7]  == n
    assert rayTraceData[8]  == l2
    assert rayTraceData[9]  == m2
    assert rayTraceData[10] == n2
    assert rayTraceData[11] == intensity
    print "zGetTrace test successful"

    print "\nTEST: zGetRefresh()"
    print "------------------"
    status = link0.zGetRefresh()
    if status == 0:
        print "Refresh successful"
    else:
        print "Refresh FAILED"

    print "\nTEST: zSetSystem()"
    print "-----------------"
    unitCode,stopSurf,rayAimingType = 0,4,0  # mm, 4th,off
    useEnvData,temp,pressure,globalRefSurf = 0,20,1,1 # off, 20C,1ATM,ref=1st surf
    systemData_s = link0.zSetSystem(unitCode,stopSurf,rayAimingType,useEnvData,
                                              temp,pressure,globalRefSurf)
    print systemData_s

    print "\nTEST: zGetSystem()"
    print "-----------------"
    systemData_g = link0.zGetSystem()
    print systemData_g

    assert systemData_s == systemData_g
    print "zSetSystem() and zGetSystem() test successful"

    print "\nTEST: zGetPupil()"
    print "------------------"
    pupil_data = dict(zip((0,1,2,3,4,5,6,7),('type','value','ENPD','ENPP',
                   'EXPD','EXPP','apodization_type','apodization_factor')))
    pupil_type = dict(zip((0,1,2,3,4,5),
            ('entrance pupil diameter','image space F/#','object space NA',
              'float by stop','paraxial working F/#','object cone angle')))
    pupil_value_type = dict(zip((0,1),("stop surface semi-diameter",
                                         "system aperture")))
    apodization_type = dict(zip((0,1,2),('none','Gaussian','Tangential')))
    # Get the pupil data
    pupilData = link0.zGetPupil()
    print "Pupil data:"
    print pupil_data[0],":",pupil_type[pupilData[0]]
    print pupil_data[1],":",pupilData[1],(pupil_value_type[0]
                             if pupilData[0]==3 else pupil_value_type[1])
    for i in range(2,6):
        print pupil_data[i],":",pupilData[i]
    print pupil_data[6],":",apodization_type[pupilData[6]]
    print pupil_data[7],":",pupilData[7]

    # Start a basic design with a new lens
    print "\nTEST: zNewLens()"
    print "----------------"
    retVal = link0.zNewLens()
    assert retVal == 0
    print "zNewLens() test successful"

    #Set (new) system parameters:
    #Get the current stop position (it should be 1, as it is a new lens)
    sysPara = link0.zGetSystem()
    # set unitCode (mm), stop-surface, ray-aiming, ... , global surface reference
    sysParaNew = link0.zSetSystem(0,sysPara[2],0,0,20,1,-1) # Set the image plane as Global ref surface

    print "\nTEST: zSetSystemAper():"
    print "-------------------"
    systemAperData_s = link0.zSetSystemAper(0,sysPara[2],25) # sysAper = 25 mm, EPD
    assert systemAperData_s[0] == 0  # Confirm aperType = EPD
    assert systemAperData_s[1] == sysPara[2]  # confirm stop surface number
    assert systemAperData_s[2] == 25  # confirm EPD value is 25 mm
    print "zSetSystemAper() test successful"

    print "\nTEST: zGetSystemAper():"
    print "-----------------------"
    systemAperData_g = link0.zGetSystemAper()
    assert systemAperData_s == systemAperData_g
    print "zGetSystemAper() test successful"

    print "\nTEST: zInsertSurface()"
    print "--------------------"
    retVal = link0.zInsertSurface(1)
    assert retVal == 0
    print "zInsertSurface() successful"

    print "\nTEST: zSetAperture()"
    print "---------------------"
    #aptInfo = link0.zSetAperture()
    pass
    #ToDo

    print "\nTEST: zGetAperture()"
    print "---------------------"
    #aptInfo = link0.zGetAperture()
    pass
    #ToDo

    print "\nTEST: zSetField()"
    print "---------------------"
    fieldData = link0.zSetField(0,0,2) # type = angle; 2 fields; rect normalization (default)
    print "fieldData: ",fieldData
    assert fieldData[0]==0; assert fieldData[1]==2;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zSetField(0,0,3,1)
    print "fieldData: ",fieldData
    assert fieldData[0]==0; assert fieldData[1]==3;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zSetField(1,0,0) # 1st field, on-axis x, on-axis y, weight = 1 (default)
    print "fieldData: ",fieldData
    assert fieldData==(0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    fieldData = link0.zSetField(2,0,5,2.0,0.5,0.5,0.5,0.5,0.5)
    print "fieldData: ",fieldData
    assert fieldData==(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5)
    fieldData = link0.zSetField(3,0,10,1.0,0.0,0.0,0.0)
    print "fieldData: ",fieldData
    assert fieldData==(0.0, 10.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    print "zSetField() test successful"

    print "\nTEST: zGetField()"
    print "---------------------"
    fieldData = link0.zGetField(0)
    print "fieldData: ",fieldData
    assert fieldData[0]==0; assert fieldData[1]==3;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zGetField(2)
    print "fieldData: ",fieldData
    assert fieldData==(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5)
    print "zGetField() successful"

    print "\nTEST: zSetFieldTuple()"
    print "---------------------"
    iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                          # vdx=vdy=vcx=vcy=van=0.0
                       (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                       (0.0,10.0))                        # field3: xf=0.0,yf=10.0
    oFieldDataTuple = link0.zSetFieldTuple(0,1,iFieldDataTuple)
    for i in range(len(iFieldDataTuple)):
        print "oFieldDataTuple, field",i,":",oFieldDataTuple[i]
        assert oFieldDataTuple[i][:len(iFieldDataTuple[i])]==iFieldDataTuple[i]
    print "zSetFieldTuple() test successful"

    print "\nTEST: zGetFieldTuple()"
    print "----------------------"
    fieldDataTuple = link0.zGetFieldTuple()
    assert fieldDataTuple==oFieldDataTuple
    print "zGetFieldTuple() test successful"

    print "\nTEST: zSetWave()"
    print "-----------------"
    wavelength1 = 0.48613270
    wavelength2 = 0.58756180
    waveData = link0.zSetWave(0,1,2)
    print "Primary wavelength number = ", waveData[0]
    print "Total number of wavelengths set = ",waveData[1]
    assert waveData[0]==1; assert waveData[1]==2
    waveData = link0.zSetWave(1,wavelength1,0.5)
    print "Wavelengths: ",waveData[0],
    assert waveData[0]==wavelength1;assert waveData[1]==0.5
    waveData = link0.zSetWave(2,wavelength2,0.5)
    print waveData[0]
    assert waveData[0]==wavelength2;assert waveData[1]==0.5
    print "zSetWave test successful"

    print "\nTEST: zGetWave()"
    print "-----------------"
    waveData = link0.zGetWave(0)
    assert waveData[0]==1;assert waveData[1]==2
    print waveData
    waveData = link0.zGetWave(1)
    assert waveData[0]==wavelength1;assert waveData[1]==0.5
    print waveData
    waveData = link0.zGetWave(2)
    assert waveData[0]==wavelength2;assert waveData[1]==0.5
    print waveData
    print "zGetWave test successful"

    print "\nTEST:zSetWaveTuple()"
    print "-------------------------"
    wavelengths = (0.48613270,0.58756180,0.65627250)
    weights = (1.0,1.0,1.0)
    iWaveDataTuple = (wavelengths,weights)
    oWaveDataTuple = link0.zSetWaveTuple(iWaveDataTuple)
    print "Output wave data tuple",oWaveDataTuple
    assert oWaveDataTuple==iWaveDataTuple
    print "zSetWaveTuple() test successful"

    print "\nTEST:zGetWaveTuple()"
    print "-------------------------"
    waveData = link0.zGetWaveTuple()
    print "Wave data tuple =",waveData
    assert oWaveDataTuple==waveData
    print "zGetWaveTuple() test successful"

    print "\nTEST: zSetPrimaryWave()"
    print "-----------------------"
    primaryWaveNumber = 2
    waveData = link0.zSetPrimaryWave(primaryWaveNumber)
    print "Primary wavelength number =", waveData[0]
    print "Total number of wavelengths =", waveData[1]
    assert waveData[0]==primaryWaveNumber
    assert waveData[1]==len(wavelengths)
    print "zSetPrimaryWave() test successful"

    print "\nTEST: zQuickFocus()"
    print "---------------------"
    #aptInfo = link0.zQuickFocus()
    pass
    #ToDo

    # Finished all tests. Perform the last test and done!
    print "\nTEST: zDDEClose()"
    print "----------------"
    status = link0.zDDEClose()
    print "Communication link 0 with ZEMAX terminated"
    status = link1.zDDEClose()
    print "Communication link 1 with ZEMAX terminated"



if __name__ == '__main__':
    import os, time
    test_PyZDDE()
