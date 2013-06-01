#-------------------------------------------------------------------------------
# Name:        pyzdde.py
# Purpose:     Python based DDE link with ZEMAX server, similar to Matlab based
#              MZDDE toolbox.
# Author:      Indranil Sinharoy
#
# Created:     08/10/2012
# Copyright:   (c) Indranil Sinharoy, 2012 - 2013
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.5
#-------------------------------------------------------------------------------
from __future__ import division
from __future__ import print_function
import win32ui
import dde
import sys
import os
from os import path
from math import pi,cos,sin
from itertools import izip
import warnings

#Import zemaxOperands
#Generally most python installations will add the current directory or the
#directory of the __main__ module to the path, it is not guaranteed. To ensure
#that the required paths are available during python search path, we add it
#explicitly. The following method of adding the path-to-the-file to python search
#path and importing the modules should be removed once something like distutils
#is used to install the module's package into the python site-packages directory.
currDir = os.path.dirname(os.path.realpath(__file__))
if currDir not in sys.path:
    sys.path.append(currDir)
import zemaxoperands as zo
import zemaxbuttons  as zb

DEBUG_PRINT_LEVEL = 0 # 0=No debug prints, but allow all essential prints
                      # 1 to 2 levels of debug print, 2 = print all

# Helper functions
def debugPrint(level,msg):
    """
    args:
        level = 0, message will definitely be printed
                1 or 2, message will be printed if level >= DEBUG_PRINT_LEVEL
        msg = string message to print
    """
    global DEBUG_PRINT_LEVEL
    if level <= DEBUG_PRINT_LEVEL:
        print("DEBUG PRINT (Level" + str(level)+ ": )" + msg)
    return

class PyZDDE(object):
    """Create an instance of PyZDDE class"""
    __chNum = -1          # channel Number
    __liveCh = 0          # number of live channels.
    __server = 0
    DDE_TIMEOUT = 3000    # Not implemented (for future), timeout (pywin32 DDE = 1 min)

    def __init__(self):
        PyZDDE.__chNum +=1   # increment ch. count when DDE ch. is instantiated.
        self.appName = "ZEMAX"+str(PyZDDE.__chNum) if PyZDDE.__chNum > 0 else "ZEMAX"
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
        debugPrint(1,"liveCh = " + str(PyZDDE.__liveCh))
        if self.appName=="ZEMAX" or PyZDDE.__liveCh==0: # do this only one time or when there is no channel
            PyZDDE.server = dde.CreateServer()
            PyZDDE.server.Create("ZCLIENT")           # Name of the client
        # Try to create individual conversations for each ZEMAX application.
        self.conversation = dde.CreateConversation(PyZDDE.server)
        try:
            self.conversation.ConnectTo(self.appName," ")
        except:
            info = sys.exc_info()
            print("Error:", info[1], ": ZEMAX may not have been started!")
            return -1
        else:
            debugPrint(1,"Zemax instance successfully connected")
            PyZDDE.__liveCh += 1 # increment the number of live channels
            self.connection = True
            #DDE_TIMEOUT = 3000 #The default timeout
            # !!! FIX: Not yet implemented.
            return 0

    def zDDEClose(self):
        """Close the DDE link with Zemax server.

        zDDEClose( ) -> Status

        Status = 0 on success.
        """
        # Close the server only if a channel was truely established and it is the last one.
        if self.connection and PyZDDE.__liveCh <=1:
            self.server.Shutdown()
            self.connection = False
            PyZDDE.__liveCh -=1  # This will become zero now. (reset)
            PyZDDE.__chNum = -1  # Reset the chNum ...
            debugPrint(2,"server shutdown")
        elif self.connection:  # if additional channels were successfully created.
            self.connection = False
            PyZDDE.__liveCh -=1
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
        PyZDDE.DDE_TIMEOUT = round(time*1000) # set time in milliseconds

    def __del__(self):
        """Destructor"""
        debugPrint(3,"Destructor called")
        self.zDDEClose()

    # ZEMAX control/query methods
    #----------------------------
    def zCloseUDOData(self,bufferCode):
        """Close the User Defined Operand (UDO) buffer, which allows the ZEMAX
        optimizer to proceed.

        zCloseUDOData(bufferCode)->retVal

        args:
            bufferCode : (integer) The buffercode is an integer value provided by
                         ZEMAX to the client that uniquely identifies the correct
                         lens.
        ret:
            retVal     : ?

        See also zGetUDOSystem and zSetUDOItem
        """
        return int(self.conversation.Request("CloseUDOData,{:d}".format(bufferCode)))

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

        See also zInsertConfig. Use zDeleteMCO() to delete a row/operand
        """
        return int(self.conversation.Request("DeleteConfig,{:d}".format(number)))

    def zDeleteMCO(self,operandNumber):
        """Deletes an existing operand (row) in the multi-configuration editor.

        zDeleteMCO(operandNumber)->newNumberOfOperands

        args:
            operandNumber        : (integer) operand number (row in the MCE) to
                                   delete.
        ret:
            newNumberOfOperands  : (integer) new number of operands.

        Note: After deleting the row, all succeeding rows (operands) are
        re-numbered.

        See also zInsertMCO. Use zDeleteConfig() to delete a column/configuration.
        """
        return int(self.conversation.Request("DeleteMCO,"+str(operandNumber)))

    def zDeleteMFO(self,operand):
        """Deletes an optimization operand (row) in the merit function editor

        zDeleteMFO(operand)->newNumOfOperands

        args:
            operand  : (integer) 1 <= operand <= number_of_operands
        ret:
            newNumOfOperands : (integer) the new number of operands

        See also zInsertMFO
        """
        return int(self.conversation.Request("DeleteMFO,{:d}".format(operand)))

    def zDeleteObject(self,surfaceNumber,objectNumber):
        """Deletes the NSC object associated with the given `objectNumber`at the
        surface associated with the `surfaceNumber`.

        zDeleteObject(sufaceNumber, objectNumber)->retVal

        args:
            surfaceNumber : (integer) surface number of Non-Sequential Component
                            surface
            objectNumber  : (integer) object number in the NSC editor.

        ret:
            retVal        : 0 if successful, -1 if it failed.

        Note: (from MZDDE) The `surfaceNumber` is 1 if the lens is fully NSC mode.
        If the command is issued when there is no more objects in, it simply
        returns 0.
        See also zInsertObject()
        """
        cmd = "DeleteObject,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            return -1
        else:
            return int(float(rs))

    def zExecuteZPLMacro(self,zplMacroCode):
        """Executes a ZPL macro present in the <data>/Macros folder.

        zExecuteZPLMacro(zplMacroCode)->status

        args:
          zplMacroCode : (string) The first 3 letters (case-sensitive) of the
                         ZPL macro present in the <data>/Macros folder.
        ret:
          status       : 0 if successfully executed the ZPL macro
                        -1 if the macro code passed is incorrect
                        error code (returned by Zemax) otherwise.

        Note:
            Limitations:
            1. Currently one can only "execute" a ZPL macro. There seems to be no
               way of getting the results of the macro execution back to the
               client application.
            2. If there are more than two macros which have the same first 3 letters
               then all of them will be executed by Zemax.
            3. The macros must be placed in the <data>/Macros folder, and the
               "ZPL" field under Preferences/Folder must be the default ZPL
               folder, i.e. <data>/Macros
        """
        status = -1
        zplMpath = path.join(self.zGetPath()[0], 'Macros')
        macroList = [f for f in os.listdir(zplMpath)
                     if f.endswith(('.zpl','.ZPL')) and f.startswith(zplMacroCode)]
        if macroList:
            zplCode = macroList[0][:3]
            status = self.zOpenWindow(zplCode,True)
        return status

    def zDeleteSurface(self,surfaceNumber):
        """Deletes an existing surface.

        zDeleteSurface(surfaceNumber)->retVal

        args:
            surfaceNumber : (integer) the surface number of the surface to be deleted
        ret:
            retVal        : 0 if successful

        Note that you cannot delete the OBJ surface (but the function still
        returns 0)
        Also see, zInsertSurface.
        """
        cmd = "DeleteSurface,{:d}".format(surfaceNumber)
        reply = self.conversation.Request(cmd)
        return int(float(reply))

    def zExportCAD(self, fileName, fileType = 1, numSpline = 32, firstSurf = 1,
                   lastSurf = -1, raysLayer = 1, lensLayer = 0, exportDummy = 0,
                   useSolids = 1, rayPattern = 0, numRays = 0, wave = 0, field = 0,
                   delVignett = 1, dummyThick = 1.00, split = 0, scatter = 0,
                   usePol = 0, config = 0):
        """Export lens data in IGES/STEP/SAT format for import into CAD programs.

        zExportCAD(exportCADdata)->status

        args:
            fileName  : (string) filename including extension (Although not necessary,
                        including full path is recommended)
            fileType  : (0, 1, 2 or 3) 0 = IGES, 1 = STEP [default], 2 = SAT, 3 = STL
            numSpline : (integer) Number of spline segments to use [default = 32]
            firstSurf : (integer) The first surface to export. In NSC mode, this
                        is the first object to export
            lastSurf  : (integer) The last surface to export. In NSC mode, this
                        is the first object to export [default = -1 i.e. image surface]
            raysLayer : (integer) Layer to place ray data on [default = 1]
            lensLayer : (integer) Layer to place lens data on [default = 0]
            exportDummy : (0 or 1) Export dummy surface. 1 = Export [default = 0]
            useSolids   : (0 or 1) Export surfaces as solids. 1 = solid surfaces [default = 1]
            rayPattern  : (0 <= rayPattern <= 7) 0 = XY fan [default], 1 = X fan, 2 = Y fan
                          3 = ring, 4 = list, 5 = none, 5 = grid, 7 = solid beams.
            numRays     : (integer) The number of rays to render [Default = 1]
            wave        : (integer) Wavelength number. 0 indicates all [Default]
            field       : (integer) The field number. 0 indicates all [Default]
            delVignett  : (0 or 1) Delete vignetted rays. 1 = delete vig. rays [Default]
            dummyThick  : (Float) Dummy surface thickness in lens units. [Default = 1.00]
            split   : (0 or 1) Split rays from NSC sources. 1 = Split sources [Default = 0]
            scatter : (0 or 1) Scatter rays from NSC sources. 1 = Scatter [Deafult = 0]
            usePol  : (0 or 1) Use polarization when tracing NSC rays. Note that
                      polarization is automatically selected if splitting is specified.
                      [Default (when splitting = 0) is 0]
            config  : (0 <= config <= n+3, where n is the total number of configurations)
                      0 = current config [Default], 1 - n for a specific configuration,
                      n+1 to export "All By File", n+2 to export "All by Layer", and
                      n+3 for "All at Once".
                      For a more detailed explanation of the configuration setting,
                      see "Export IGES/SAT.STEP Solid" in the manual.
        rets:
            status : (string) the string "Exporting filename" or "BUSY!" (see
                     description below)

        There is a complexity in using this feature via DDE. The export of lens
        data may take a long time relative to the timeout interval of the DDE
        communication. Therefore, calling this data item will cause ZEMAX to launch
        an independent thread to process the request. Once the thread is launched,
        the return value is the string "Exporting filename". However, the actual
        file may take many seconds or minutes to be ready to use. To verify that
        the export is complete and the file is ready, use the zExportCheck() function.
        zExportCheck() will return 1 if the export is still running, or 0 if it
        has completed. Generally, the zExportCheck() function call will need to
        be placed in a loop which executes until zExportCheck() returns 0.

        A typical loop test in Python code might look like this (assuming `ddelink`
        is an instance of PyZDDE):

        # check if the export is done
        still_working = True
        while(still_working):
            # Delay for 200 milliseconds
            time.sleep(.2)
            status = ddelink.zExportCheck()
            if status == 1:  # still running
                pass
            elif status == 0: # Done exporting
                still_working = False

        Note: Zemax cannot export some NSC objects such as slide. In such cases
        the unexportable objects are ignored.
        """
        #Determine last surface/object depending upon zemax mode
        if lastSurf == -1:
            zmxMode = self.zGetMode()
            if zmxMode[0] != 1:
                lastSurf = self.zGetSystem()[0]
            else:
                lastSurf = self.zGetNSCData(1,0)
        args = [str(arg) for arg in ("ExportCAD", fileName, fileType, numSpline,
                                     firstSurf, lastSurf, raysLayer, lensLayer,
                                     exportDummy, useSolids, rayPattern, numRays,
                                     wave, field, delVignett, dummyThick, split,
                                     scatter, usePol, config)]
        cmd = ",".join(args)
        reply = self.conversation.Request(cmd)
        return str(reply)

    def zExportCheck(self):
        """Used to indicate the status of the last executed zExportCAD() command.

        zExportCheck()->status

        args:
            None
        ret:
            status : (integer) 0 = last CAD export completed
                               1 = last CAD export in progress
        """
        return int(self.conversation.Request('ExportCheck'))

    def zFindLabel(self,label):
        """Returns the surface that has the integer label associated with the
        specified surface.

        zFindLabel(label)->surfaceNumber

        args:
            label   :  (integer) label
        ret:
            surfaceNumber : (integer) surface number of surface associated with
                            the `label`. It returns -1 if no surface has the
                            specified label.

        See also zSetLabel, zGetLabel()
        """
        reply = self.conversation.Request("FindLabel,{:d}".format(label))
        return int(float(reply))

    def zGetAddress(self,addressLineNumber):
        """Extract the address line number indicated by `addressLineNumber`

        zGetAddress(addressLineNumber)->addressLine

        args:
            addressLineNumber  : (integer) line number of address to get
        ret:
            addressLine        : (string) address line
        """
        reply = self.conversation.Request("GetAddress,{:d}"
                                          .format(addressLineNumber))
        return str(reply)

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
                           see "Aperture type & other aperture controls" for more details.
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

    def zGetApodization(self,px,py):
        """Computes the intensity apodization of a ray from the apodization
        type and value.

        zGetApodication(px,py)->intensityApodization

        args:
            px,py  : (float) normalized pupil coordinates.
        ret:
            intensityApodization : (float) intensity apodization
        """
        reply = self.conversation.Request("GetApodization,{:1.20g},{:1.20g}"
                                          .format(px,py))
        return float(reply)

    def zGetAspect(self,filename=None):
        """Returns the graphic display aspect ratio and the width or height of the
        printed page in current lens units.

        zGetAspect([filename])->(aspect,side)

        args:
            filename : name of the temporary file associated with the window
                       being created or updated. If the temporary file is left
                       off, then the default aspect ratio and width (or height)
                       is returned.
        rets:
            aspect : aspect ratio (height/width)
            side   : width if aspect <= 1; height if aspect > 1. (in lens units)
        """
        if filename == None:
            cmd = "GetAspect"
        else:
            cmd = "GetAspect,{}".format(filename)
        reply = self.conversation.Request(cmd)
        rs = reply.split(",")
        aspectSide = tuple([float(elem) for elem in rs])
        return aspectSide

    def zGetBuffer(self,n,tempFileName):
        """Retrieve ZEMAX DDE client specific data from a window being updated.

        zGetBuffer(n,tempFileName)->bufferData

        args:
          n            : (integer, 0<=n<=15) the buffer number
          tempFileName : name of the temporary file associated with the window
                         being updated. The tempfile name is passed to the client
                         when ZEMAX calls the client; see the discussion
                         "How ZEMAX calls the client" in the Zemax manual for
                         details.
        ret:
          bufferData   : (string)

        Note each window may have it's own buffer data, and ZEMAX uses the
        filename to identify the window for which the buffer contents are required.

        See also zSetBuffer.
        """
        cmd = "GetBuffer,{:d},{}".format(n,tempFileName)
        reply = self.conversation.Request(cmd)
        return str(reply.rstrip())
        # !!!FIX what is the proper return for this command?

    def zGetComment(self,surfaceNumber):
        """Returns the surface comment, if any, associated with the surface

        zGetComment(surfaceNumber)->comment

        args:
          surfaceNumber: (integer) the surface number
        ret:
          comment      : (string) the comment, if any, associated with the surface

        """
        reply = self.conversation.Request("GetComment,{:d}".format(surfaceNumber))
        return str(reply.rstrip())

    def zGetConfig(self):
        """Returns the current configuration number (selected column in the MCE),
        the number of configurations (number of columns), and the number of
        multiple configuration operands (number of rows).

        zGetConfig()->(currentConfig, numberOfConfigs, numberOfMutiConfigOper)

        args:
          none
        ret:
          3-tuple containing the following elements:
           currentConfig          : current configuration (column) number in MCE
           numberOfConfigs        : number of configs (columns)
           numberOfMutiConfigOper : number of multi config operands (rows)

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
        # !!! FIX: Should this function return "0" when the MCE is empty, just
        # like what is done for the zGetNSCData() function?
        return tuple([int(elem) for elem in rs])

    def zGetDate(self):
        """Request current date from the ZEMAX DDE server.

        zGetDate()->date

        args:
          None
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
        cmd="GetExtra,{sn:d},{cn:d}".format(sn=surfaceNumber,cn=columnNumber)
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

    def zGetGlass(self,surfaceNumber):
        """Returns some data about the glass on any surface.

        zGetGlass(surfaceNumber)->glassInfo

        args:
          surfaceNumber : (integer) surface number
        ret:
          glassInfo   : 3-tuple containing the name, nd,vd, dpgf if there is a
                        valid glass associated with the surface, else `None`

        Note: If the specified surface is not valid, is not made of glass, or is
        gradient index, the returned string is empty. This data may be meaningless
        for glasses defined only outside of the FdC band.
        """
        reply = self.conversation.Request("GetGlass,{:d}".format(surfaceNumber))
        rs = reply.split(',')
        if len(rs) > 1:
            glassInfo = tuple([str(rs[i]) if i == 0 else float(rs[i])
                                                      for i in range(len(rs))])
        else:
            glassInfo = None
        return glassInfo

    def zGetGlobalMatrix(self,surfaceNumber):
        """Returns the the matrix required to convert any local coordinates (such
        as from a ray trace) into global coordinates.

        zGetGlobalMatrix(surfaceNumber)->globalMatrix

        args:
          surfaceNumber : (integer) surface number
        ret:
          globalMatrix  : is a 9-tuple, if successful  = (R11,R12,R13,
                                                          R21,R22,R23,
                                                          R31,R32,R33,
                                                          Xo, Yo , Zo)
                          it returns -1, if bad command.

        For details on the global coordinate matrix, see "Global Coordinate Reference
        Surface" in the Zemax manual.
        """
        cmd = "GetGlobalMatrix,{:d}".format(surfaceNumber)
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        globalMatrix = tuple([float(elem) for elem in rs.split(',')])
        return globalMatrix

    def zGetIndex(self,surfaceNumber):
        """Returns the index of refraction data for any surface.

        zGetIndex(surfaceNumber)->indexTuple

        args:
          surfaceNumber : (integer) surface number
        ret:
          indexTuple    : tuple of (real) index of refraction values,
                          defined for each wavelength. (n1,n2,n3,...)
                          If surface number is not valid,
        """
        reply = self.conversation.Request("GetIndex,{:d}".format(surfaceNumber))
        rs = reply.split(",")
        indexTuple = [float(rs[i]) for i in range(len(rs))]
        return tuple(indexTuple)


    def zGetLabel(self,surfaceNumber):
        """This command retrieves the integer label assicuated with the specified
        surface. Labels are be retained by ZEMAX as surfaces are inserted or deleted
        around the target surface.

        zGetLabel(surfaceNumber)->label

        args:
          surfaceNumber : (integer) the surface number
        rets:
          label         : (integer) the integer label

        See also zSetLabel, zFindLabel
        """
        reply = self.conversation.Request("GetLabel,{:d}".format(surfaceNumber))
        return int(float(reply.rstrip()))

    def zGetMetaFile(self,metaFileName,analysisType,settingsFileName,flag):
        """Creates a windows Metafile of any ZEMAX graphical analysis plot.

           zMetaFile(metaFilename, analysisType, settingsFileName, flag)->retVal

           args:
            metaFileName : name of the file to be created including the full path,
                           name, and extension for the metafile.
            analysisType : 3 letter case-sensitive label that indicates the
                           type of the analysis to be performed. They are identical
                           to those used for the button bar in Zemax. The labels
                           are case sensitive. If no label is provided or recognized,
                           a 3D Layout plot will be generated.
            settingsFileName : If a valid file name is used for the "settingsFileName",
                               ZEMAX will use or save the settings used to compute
                               the metafile graphic, depending upon the value of
                               the flag parameter.
            flag        :  0 = default settings used for the graphic
                           1 = settings provided in the settings file, if valid,
                               else default settings used
                           2 = settings provided in the settings file, if valid,
                               will be used and the settings box for the requested
                               feature will be displayed. After the user makes any
                               changes to the settings the graphic will then be
                               generated using the new settings.
                           Please see the ZEMAX manual for more details.
           retVal:
            0      : Success
            -1     : Metafile could not be saved (Zemax may not have received
                     a full path name or extention).
            -998   : Command timed out

        Notes: No matter what the flag value is, if a valid file name is provided
        for the settingsfilename, the settings used will be written to the settings
        file, overwriting any data in the file.

        Example: zGetMetaFile("C:\myGraphicfile.EMF",'Lay',"None",0)

        See also zGetTextFile, zOpenWindow.
        """
        retVal = -1
        #Check if the file path is valid and has extension
        if path.isabs(metaFileName) and path.splitext(metaFileName)[1]!='':
            cmd = 'GetMetaFile,"{tF}",{aT},"{sF}",{fl:d}'.format(tF=metaFileName,
                                    aT=analysisType,sF=settingsFileName,fl=flag)
            reply = self.conversation.Request(cmd)
            if 'OK' in reply.split():
                retVal = 0
        return retVal

    def zGetMode(self):
        """Returns the mode (Sequential, Non-sequential or Mixed) of the current
        lens in the DDE server. For the purpose of this function, "Sequential"
        implies that there are no non-sequential surfaces in the LDE.

        zGetMode()->zmxModeInformation

        args:
            None
        ret:
            zmxModeInformation is a 2-tuple, with the second element, `nscSurfNums`
            also a tuple. I.e. zmxModeInformation = (mode,nscSurfNums)
            mode : (integer) 0 = Sequential, 1 = Non-sequential, 2 = Mixed mode
            nscSurfNums : (tuple of integers) the surfaces (in mixed mode) that
                          are non-sequential. In Non-sequential mode and in purely
                          sequential mode, this tuple is empty (of length 0).

        Note: This function is not specified in the Zemax manual
        """
        nscSurfNums = []
        nscData = self.zGetNSCData(1,0)
        if nscData > 0: # Non-sequential mode
            mode = 1
        else:          # Not Non-sequential mode
            numSurf = self.zGetSystem()[0]
            for i in range(1,numSurf+1):
                surfType = self.zGetSurfaceData(i,0)
                if surfType == 'NONSEQCO':
                    nscSurfNums.append(i)
            if len(nscSurfNums) > 0:
                mode = 2  # mixed mode
            else:
                mode = 0  # sequential
        return (mode,tuple(nscSurfNums))

    def zGetMulticon(self,config,row):
        """Extract data from the multi-configuration editor.

        zGetMulticon(config,row)->multiConData

        args:
            config : (integer) configuration number (column)
            row    : (integer) operand
        ret:
            multiConData is a tuple whose elements are dependent on the value of
            `config`

            If `config` > 0, then the elements of multiConData are:
                (value,num_config,num_row,status,pickuprow,pickupconfig,scale,offset)

              The status integer is 0 for fixed, 1 for variable, 2 for pickup, and 3
              for thermal pickup. If status is 2 or 3, the pickuprow and pickupconfig
              values indicate the source data for the pickup solve.

            If `config` = 0, then the elements of multiConData are:
                (operand_type,number1,number2,number3)

        See also zSetMulticon.
        """
        cmd = "GetMulticon,{config:d},{row:d}".format(config=config,row=row)
        reply = self.conversation.Request(cmd)
        if config: # if config > 0
            rs = reply.split(",")
            if '' in rs: # if the MCE is "empty"
                rs[rs.index('')] = '0'
            multiConData = [float(rs[i]) if (i == 0 or i == 6 or i== 7) else int(rs[i])
                                                 for i in range(len(rs))]
        else: # if config == 0
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0,rs[0])
        return tuple(multiConData)

    def zGetName(self):
        """Returns the name of the lens.

        zGetName()->lensName

        args:
            None
        ret:
            lensName  : (string) name of the current lens (as entered on the
                        General data dialog box) in the DDE server
        """
        reply = self.conversation.Request('GetName')
        return str(reply.rstrip())

    def zGetNSCData(self,surfaceNumber,code):
        """Returns the data for NSC groups.

        zGetNSCData(surface,code)->nscData

        args:
            surfaceNumber  : (integer) surface number of the NSC group. Use 1 if
                             the program mode is Non-Sequential.
            code           : Currently only code = 0 is supported, in which case
                             the returned data is the number of objects in the
                             NSC group
        rets:
            nscData  : the number of objects in the NSC group if the command
                       was successful (valid).
                       -1 if it was a bad commnad (generally if the `surface` is
                       not a non-sequential surface)

        Note: this function returns 1 if the only object in the NSC editor is a
        "Null Object".
        """
        cmd = "GetNSCData,{:d},{:d}".format(surfaceNumber,code)
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscData = -1
        else:
            nscData = int(float(rs))
            if nscData == 1:
                nscObjType = self.zGetNSCObjectData(surfaceNumber,1,0)
                if nscObjType == 'NSC_NULL': # the NSC editor is actually empty
                    nscData = 0
        return nscData

    def zGetNSCMatrix(self,surfaceNumber,objectNumber):
        """Returns a tuple containing the rotation and position matrices relative
        to the NSC surface origin.

        zGetNSCMatrix(surfaceNumber,objectNumber)->nscMatrix

        args:
            surfaceNumber : (integer) surface number of the NSC group. Use 1 if
                            the program mode is Non-Sequential.
            objectNumber  : (integer) the NSC ojbect number
        ret:
            nscMatrix     : is a 9-tuple, if successful  = (R11,R12,R13,
                                                            R21,R22,R23,
                                                            R31,R32,R33,
                                                            Xo, Yo , Zo)
                            it returns -1, if bad command.
        """
        cmd = "GetNSCMatrix,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscMatrix = -1
        else:
            nscMatrix = tuple([float(elem) for elem in rs.split(',')])
        return nscMatrix

    def zGetNSCObjectData(self,surfaceNumber,objectNumber,code):
        """Returns the various data for NSC objects.

        zGetNSCOjbect(surfaceNumber,objectNumber,code)->nscObjectData

        args:
            surfaceNumber : (integer) surface number of the NSC group. Use 1 if
                            the program mode is Non-Sequential.
            objectNumber  : (integer) the NSC ojbect number
            code          : (integer) see the nscObjectData returned table
        rets:
            nscObjectData : nscObjectData as per the table below if successful
                            else -1
        ------------------------------------------------------------------------
        Code - Data returned by GetNSCObjectData
        ------------------------------------------------------------------------
          0  - Object type name. (string)
          1  - Comment, which also defines the file name if the object is defined
               by a file. (string)
          2  - Color. (integer)
          5  - Reference object number. (integer)
          6  - Inside of object number. (integer)
        The following codes set values on the Type tab of the Object Properties dialog.
          3  - 1 if object uses a user defined aperture file, 0 otherwise. (integer)
          4  - User defined aperture file name, if any. (string)
         29  - Gets the "Use Pixel Interpolation" checkbox. (1 = checked, 0 = unchecked)
        The following codes set values on the Sources tab of the Object Properties dialog.
        101  - Gets the source object random polarization. (1 = checked, 0 = unchecked)
        102  - Gets the source object reverse rays option. (1 = checked, 0 for unchecked)
        103  - Gets the source object Jones X value.
        104  - Gets the source object Jones Y value.
        105  - Gets the source object Phase X value.
        106  - Gets the source object Phase Y value.
        107  - Gets the source object initial phase in degrees value.
        108  - Gets the source object coherence length value.
        109  - Gets the source object pre-propagation value.
        110  - Gets the source object sampling method; (0 = random, 1 = Sobol sampling)
        111  - Gets the source object bulk scatter method; (0 = many, 1 = once, 2 = never)
        The following codes set values on the Bulk Scatter tab of the Object
        Properties dialog.
        202  - Gets the Mean Path value.
        203  - Gets the Angle value.
        211-226 - Gets the DLL parameter 1-16, respectively.
        ------------------------------------------------------------------------
        """
        str_codes = (0,1,4)
        int_codes = (2,3,5,6,29,101,102,110,111)
        cmd = ("GetNSCObjectData,{:d},{:d},{:d}"
              .format(surfaceNumber,objectNumber,code))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscObjectData = -1
        else:
            if code in str_codes:
                nscObjectData = str(rs)
            elif code in int_codes:
                nscObjectData = int(float(rs))
            else:
                nscObjectData = float(rs)
        return nscObjectData

    def zGetNSCObjectFaceData(self,surfNumber,objNumber,faceNumber,code):
        """Returns the various data for NSC object faces.

        zGetNSCObjectFaceData(surfNumber,objNumber,faceNumber,code)->nscObjFaceData

        args:
          surfNumber   : (integer) surface number. Use 1 if the program mode is
                         Non-Sequential.
          objNumber    : (integer) object number
          faceNumber   : (integer) face number
          code         : (integer) code (see below)
        ret:
          nscObjFaceData  : data for NSC object faces (see the table for the
                            particular type of data) if successful, else -1
        ------------------------------------------------------------------------
        Code     Data returned by GetNSCObjectFaceData
        ------------------------------------------------------------------------
         10   -  Coating name. (string)
         20   -  Scatter code. (0 = None, 1 = Lambertian, 2 = Gaussian,
                 3 = ABg, and 4 = user defined.) (integer)
         21   -  Scatter fraction. (double)
         22   -  Number of rays to scatter. (integer)
         23   -  Gaussian scatter sigma. (double)
         24   -  Face is setting;(0 = object default, 1 = reflective,
                 2 = absorbing.) (integer)
         30   -  ABg scatter profile name for reflection. (string)
         31   -  ABg scatter profile name for transmission. (string)
         40   -  User Defined Scatter DLL name. (string)
         41-46 - User Defined Scatter Parameter 1 - 6. (double)
         60   -  User Defined Scatter data file name. (string)
        ------------------------------------------------------------------------
        See also zSetNSCObjectFaceData
        """
        str_codes = (10,30,31,40,60)
        int_codes = (20,22,24)
        cmd = ("GetNSCObjectFaceData,{:d},{:d},{:d},{:d}"
              .format(surfNumber,objNumber,faceNumber,code))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscObjFaceData = -1
        else:
            if code in str_codes:
                nscObjFaceData = str(rs)
            elif code in int_codes:
                nscObjFaceData = int(float(rs))
            else:
                nscObjFaceData = float(rs)
        return nscObjFaceData

    def zGetNSCParameter(self,surfNumber,objNumber,parameterNumber):
        """Returns the parameter data for NSC objects.

        zGetNSCParameter(surfNumber,objNumber,parameterNumber)->nscParaVal

        args:
          surfNumber      : (integer) surface number. Use 1 if
                            the program mode is Non-Sequential.
          objNumber       : (integer) object number
          parameterNumber : (integer) parameter number
        ret:
          nscParaVal     : (float) parameter value

        See also zSetNSCParameter
        """
        cmd = ("GetNSCParameter,{:d},{:d},{:d}"
              .format(surfNumber,objNumber,parameterNumber))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zGetNSCPosition(self,surfNumber,objectNumber):
        """Returns the position data for NSC objects.

        zGetNSCPosition(surfNumber,objectNumber)->nscPosData

        arg:
          surfNumber   : (integer) surface number. Use 1 if
                         the program mode is Non-Sequential.
          objectNumber:  (integer) object number
        ret:
          nscPosData is a 7-tuple containing x,y,z,tilt-x,tilt-y,tilt-z,material

        See also zSetNSCPosition
        """
        cmd = ("GetNSCPosition,{:d},{:d}".format(surfNumber,objectNumber))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPosData = -1
        else:
            nscPosData = tuple([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                    for i in range(len(rs))])
        return nscPosData

    def zGetNSCProperty(self,surfaceNumber,objectNumber,faceNumber,code):
        """Returns a numeric or string value from the property pages of objects
        defined in the non-sequential components editor. It mimics the ZPL
        function NPRO.

        zGetNSCProperty(surfaceNumber,objectNumber,faceNumber,code)->nscPropData

        args:
          surfaceNumber  : (integer) surface number. Use 1 if the program mode is
                           Non-Sequential.
          objectNumber   : (integer) object number
          faceNumber     : (integer) face number
          code           : (integer) code used to identify specific types of NSC
                           object data (see below)
        ret:
          nscPropData  :  string or numeric (see below)
        ------------------------------------------------------------------------
        CODE -             PROPERTY
        ------------------------------------------------------------------------
        The following codes get values from the NSC Editor.
          1 - Gets the object comment. (string)
          2 - Gets the reference object number. (integer)
          3 - Gets the "inside of" object number. (integer)
          4 - Gets the object material. (string)
        The following codes get values from the Type tab of the Object Properties
        dialog.
          0 - Gets the object type. The value should be the name of the object,
              such as "NSC_SLEN" for the standard lens. The names for each object
              type are listed in the Prescription Report for each object type in
              the NSC editor. All NSC object names start with "NSC_". (string)
         13 - Gets User Defined Aperture, (1 = checked, 0 = unchecked)
         14 - Gets the User Defined Aperture file name. (string)
         15 - Gets the "Use Global XYZ Rotation Order" checkbox, (1 = checked,
              0 = unchecked). (integer)
         16 - Gets the "Rays Ignore This Object" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
         17 - Gets the "Object Is A Detector" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
         18 - Gets the "Consider Objects" list. The argument should be a string
              listing the object numbers to consider delimited by spaces, such
              as "2 5 14". (string)
         19 - Gets the "Ignore Objects" list. The argument should be a string
              listing the object numbers to ignore delimited by spaces, such as
              "1 3 7". (string)
         20 - Gets the "Use Pixel Interpolation" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
        The following codes get values from the Coat/Scatter tab of the Object
        Properties dialog.
          5 - Gets the coating name for the specified face. (string)
          6 - Gets the profile name for the specified face. (string)
          7 - Gets the scatter mode for the specified face, (0 = none,
              1 = Lambertian, 2 = Gaussian, 3 = ABg, 4 = User Defined.)
          8 - Gets the scatter fraction for the specified face. (float)
          9 - Gets the number of scatter rays for the specified face. (integer)
         10 - Gets the Gaussian sigma for the specified face. (float)
         11 - Gets the reflect ABg data name for the specified face. (string)
         12 - Gets the transmit ABg data name for the specified face. (string)
         27 - Gets the name of the user defined scattering DLL. (string)
         28 - Gets the name of the user defined scattering data file. (string)
        21-26 - Gets parameter values on the user defined scattering DLL. (float)
         29 - Gets the "Face Is" property for the specified face. (0 = "Object
              Default", 1 = "Reflective", 2 = "Absorbing")
        The following codes get values from the Bulk Scattering tab of the Object
        Properties dialog.
         81 - Gets the "Model" value on the bulk scattering tab. (0 = "No Bulk
              Scattering", 1 = "Angle Scattering", 2 = "DLL Defined Scattering")
         82 - Gets the mean free path to use for bulk scattering.
         83 - Gets the angle to use for bulk scattering.
         84 - Gets the name of the DLL to use for bulk scattering.
         85 - Gets the parameter value to pass to the DLL, where the face value
              is used to specify which parameter is being defined. The first
              parameter is 1, the second is 2, etc. (float)
         86 - Gets the wavelength shift string. (string)
        The following codes get values from the Diffraction tab of the Object
        Properties dialog.
         91 - Gets the "Split" value on the diffraction tab. (0 = "Don't Split
              By Order", 1 = "Split By Table Below", 2 = "Split By DLL Function")
         92 - Gets the name of the DLL to use for diffraction splitting. (string)
         93 - Gets the Start Order value. (float)
         94 - Gets the Stop Order value. (float)
         95 - Gets  the  parameter  values  on  the  diffraction  tab.  These
              are  the  parameters passed to the diffraction splitting DLL as
              well as the order efficiency values used by "split by table below"
              option. The face value is used to specify which parameter is being
              defined. The first parameter is 1, the second is 2, etc. (float)
        The following codes get values from the Sources tab of the Object
        Properties dialog.
        101 - Gets the source object random polarization.(1=checked,0=unchecked)
        102 - Gets the source object reverse rays option.(1=checked,0=unchecked)
        103 - Gets the source object Jones X value.
        104 - Gets the source object Jones Y value.
        105 - Gets the source object Phase X value.
        106 - Gets the source object Phase Y value.
        107 - Gets the source object initial phase in degrees value.
        108 - Gets the source object coherence length value.
        109 - Gets the source object pre-propagation value.
        110 - Gets the source object sampling method; (0=random,1=Sobol sampling)
        111 - Gets the source object bulk scatter method; (0=many,1=once,2=never)
        112 - Gets the array mode; (0 = none, 1 = rectangular, 2 = circular,
              3 = hexapolar, 4 = hexagonal)
        113 - Gets the source color mode. For a complete list of the available
              modes, see "Defining the color and spectral content of sources"
              in the Zemax manual. The source color modes are numbered starting
              with 0 for the System Wavelengths, and then from 1 through the last
              model listed in the dialog box control. (integer)
        114-116 - Gets the number of spectrum steps, start wavelength, and end
                  wavelength, respectively. (float)
        117 - Gets the name of the spectrum file. (string)
        161-162 - Gets the array mode integer arguments 1 and 2.
        165-166 - Gets the array mode double precision arguments 1 and 2.
        181-183 - Gets the source color mode arguments, for example, the XYZ
                  values of the Tristimulus. (float)
        The following codes get values from the Grin tab of the Object Properties
        dialog.
        121 - Gets the "Use DLL Defined Grin Media" checkbox. (1 = checked, 0 =
              unchecked) (integer)
        122 - Gets the Maximum Step Size value. (float)
        123 - Gets the DLL name. (string)
        124 - Gets the Grin DLL parameters. These are the parameters passed to
              the DLL. The face value is used to specify which parameter is being
              defined. The first parameter is 1, the second is 2, etc. (float)
        The following codes get values from the Draw tab of the Object Properties
        dialog.
        141 - Gets the do not draw object checkbox.(1 = checked, 0 = unchecked)
        142 - Gets the object opacity. (0 = 100%, 1 = 90%, 2 = 80%, etc.)
        The following codes get values from the Scatter To tab of the Object
        Properties dialog.
        151 - Gets the scatter to method. (0 = scatter to list, 1 = importance
              sampling)
        152 - Gets  the  Importance  Sampling  target  data.  The  argument
              should be a string listing the ray number, the object number, the
              size, & the limit value, all separated by spaces Here is a sample
              syntax to set the Importance Sampling data for ray 3, object 6,
              size 3.5, and limit 0.6:"3 6 3.5 0.6". (string)
        153 - Gets the "Scatter To List" values. The argument should be a string
              listing the object numbers to scatter to delimited by spaces, such
              as "4 6 19". (string)
        The following codes get values from the Birefringence tab of the Object
        Properties dialog.
        171 - Gets the Birefringent Media checkbox. (0 = unchecked, 1 = checked)
        172 - Gets the Birefringent Media Mode. (0 = Trace ordinary and
              extraordinary rays, 1 = Trace only ordinary rays, 2 = Trace only
              extraordinary rays, and 3 = Waveplate mode) (integer)
        173 - Gets the Birefringent Media Reflections status. (0 = Trace
              reflected and refracted rays, 1 = Trace only refracted rays, and
              2 = Trace only reflected rays)
        174-176 - Gets the Ax, Ay, and Az values. (float)
        177 - Gets the Axis Length. (float)
        200 - Gets the index of refraction of an object.
        201-203 - Gets the nd (201), vd (202), and dpgf (203) parameters of an
                  object using a model glass.
        ------------------------------------------------------------------------
        See also zSetNSCProperty
        """
        cmd = ("GetNSCProperty,{:d},{:d},{:d},{:d}"
                .format(surfaceNumber,objectNumber,code,faceNumber))
        reply = self.conversation.Request(cmd)
        nscPropData = process_get_set_NSCProperty(code,reply)
        return nscPropData

    def zGetNSCSettings(self):
        """Returns the maximum number of intersections, segments, nesting level,
        minimum absolute intensity, minimum relative intensity, glue distance,
        miss ray distance, and ignore errors flag used for NSC ray tracing.

        zGetNSCSettings()->nscSettingsData

        args:
          None
        rets:
          nscSettingsData is an 8-tuple with the following elements
          maxInt     : (integer) maximum number of intersections
          maxSeg     : (integer) maximum number of segments
          maxNest    : (integer) maximum nesting level
          minAbsI    : (float) minimum absolute intensity
          minRelI    : (float) minimum relative intensity
          glueDist   : (float) glue distance
          missRayLen : (float) miss ray distance
          ignoreErr  : (integer) 1 if true, 0 if false

        See also zSetNSCSettings
        """
        reply = str(self.conversation.Request('GetNSCSettings'))
        rs = reply.rsplit(",")
        nscSettingsData = [float(rs[i]) if i in (3,4,5,6) else int(float(rs[i]))
                                                        for i in range(len(rs))]
        return tuple(nscSettingsData)

    def zGetOperand(self,row,column):
        """Returns the operand data from the Merit Function Editor.

        zGetOperand(row,column)-> operandData

        args:
          row   : (integer) row operand number in the MFE
          column : (integer) column
        ret:
          operandData : integer, float or string depending upon column  (see
                        table below) if successful, else -1.
                        -----------------------------------------------
                        Column        Returned operand data
                        -----------------------------------------------
                         1               operand type (string)
                         2               int1 (integer)
                         3               int2 (integer)
                         4-7             data1-data4 (float)
                         8               target (float)
                         9               weight (float)
                         10              contribution (float)
                         12-13           data5-data6 (float)
                        -----------------------------------------------
        Note: To update the merit function prior to calling zGetOperand function,
              use the zOptimize() function with the number of cycles set to -1.

        See also zSetOperand and zOptimize.
        """
        cmd = "GetOperand,{:d},{:d}".format(row, column)
        reply = self.conversation.Request(cmd)
        return process_get_set_Operand(column, reply)

    def zGetPath(self):
        """Returns path name to <data> folder and default lenses folder.

        zGetPath()->(pathToDataFolder,pathToDefaultLensFolder)

        args:
            None
        ret:
            pathToDataFolder : (string) full path to the <data> folder
            pathToDefaultLensFolder : (string) full path to the default folder for
                                      lenses.
        """
        reply = str(self.conversation.Request('GetPath'))
        rs = str(reply.rsplit()[0])
        return tuple(rs.split(','))

    def zGetPolState(self):
        """Returns the default polarization state set by the user.

        zGetPolState()->polStateData

        args:
          None
        ret:
          polStateData is 5-tuple containing the following elements
          nlsPolarized : (integer) if nlsPolarized > 0, then default polarization
                         state is unpolarized.
          Ex           : (float) normalized electric field magnitude in x direction
          Ey           : (float) normalized electric field magnitude in y direction
          Phax         : (float) relative phase in x direction in degrees
          Phay         : (float) relative phase in y direction in degrees

        Note: The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
              values are accepted.

        See also zSetPolState.
        """
        reply = self.conversation.Request("GetPolState")
        rs = reply.rsplit(",")
        polStateData = [int(float(elem)) if i==0 else float(elem)
                                       for i,elem in enumerate(rs[:-1])]
        return tuple(polStateData)

    def zGetPolTrace(self,waveNum,mode,surf,hx,hy,px,py,Ex,Ey,Phax,Phay):
        """Trace a single polarized ray through the current lens in the ZEMAX
        DDE server. If Ex, Ey, Phax, Phay are all zero, Zemax will trace two
        orthogonal rays and the resulting transmitted intensity will be averaged.

        zGetPolTrace(waveNum,mode,surf,hx,hy,px,py,Ex,Ey,Phax,Phay)->rayPolTraceData

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
          Ex      : normalized electric field magnitude in x direction
          Ey      : normalized electric field magnitude in y direction
          Phax    : relative phase in x direction in degrees
          Phay    : relative phase in y direction in degrees

        ret:
          rayPolTraceData : rayPolTraceData is a 8-tuple or 2-tuple (depending
          on polarized or unpolarized rays) containing the following elements:

          For polarized rays --
              error       : 0, if the ray traced successfully
                            +ve number indicates that the ray missed the surface
                            -ve number indicates that the ray total internal
                            reflected (TIR) at the surface given by the absolute
                            value of the errorCode number.
              intensity   : the transmitted intensity of the ray. It is always
                            normalized to an input electric field intensity of
                            unity. The transmitted intensity accounts for surface,
                            thin film, and bulk absorption effects, but does not
                            consider whether or not the ray was vignetted.
              Exr,Eyr,Ezr : real parts of the electric field components
              Exi,Eyi,Ezi : imaginary parts of the electric field components

          For unpolarized rays --
              error       : (see above)
              intensity   : (see above)

        Example: To trace the real unpolarized marginal ray to the image surface
                 at wavelength 2, the function would be:
                 zGetPolTrace(2,0,-1,0.0,0.0,0.0,1.0,0,0,0,0)

        Note:
        1. The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
           values are accepted.
        2. There is an important exception to the above rule -- If Ex, Ey, Phax,
           Phay are all zero, Zemax will trace two orthogonal rays and the resul-
           ting transmitted intensity will be averaged.
        3. Always check to verify the ray data is valid (check the error) before
           using the rest of the data in the tuple.
        4. Use of zGetPolTrace() has significant overhead as only one ray per DDE
           call is traced. Please refer to the ZEMAX manual for more details.
           Also, if a large number of rays are to be traced, see the section
           "Tracing large number of rays" in the ZEMAX manual.

        See also zGetPolTraceDirect, zGetTrace, zGetTraceDirect
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f}".format(px=px,py=py)
        args4 = "{Ex:1.4f},{Ey:1.4f}".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = "GetPolTrace," + args1 + args2 + args3 + args4 + args5
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        rayPolTraceData = tuple([int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)])
        return rayPolTraceData

    def zGetPolTraceDirect(self,waveNum,mode,
                           startSurf,stopSurf,
                           x,y,z,l,m,n,Ex,Ey,Phax,Phay):
        """Trace a single polarized ray through the current lens in the ZEMAX
        DDE server while providing a more direct access to the ZEMAX ray tracing
        engine than zGetPolTrace. If Ex, Ey, Phax, Phay are all zero, Zemax will
        trace two orthogonal rays and the resulting transmitted intensity will be
        averaged.

        zGetPolTrace(waveNum,mode,startSurf,stopSurf,x,y,z,
                     l,m,n,Ex,Ey,Phax,Phay)->rayPolTraceData

        args:
          waveNum  : wavelength number as in the wavelength data editor
          mode     : 0 = real, 1 = paraxial
          startSurf: surface to trace the ray from.
          stopSurf : last surface to trace the polarized ray to.
          x,y,z,   : coordinates of the ray at the starting surface
          l,m,n    : the direction cosines to the entrance pupil aim point for
                     the x-, y-, z- direction cosines respectively
          Ex       : normalized electric field magnitude in x direction
          Ey       : normalized electric field magnitude in y direction
          Phax     : relative phase in x direction in degrees
          Phay     : relative phase in y direction in degrees

        ret:
          rayPolTraceData : rayPolTraceData is a 8-tuple or 2-tuple (depending
          on polarized or unpolarized rays) containing the following elements:

          For polarized rays --
              error       : 0, if the ray traced successfully
                            +ve number indicates that the ray missed the surface
                            -ve number indicates that the ray total internal
                            reflected (TIR) at the surface given by the absolute
                            value of the errorCode number.
              intensity   : the transmitted intensity of the ray. It is always
                            normalized to an input electric field intensity of
                            unity. The transmitted intensity accounts for surface,
                            thin film, and bulk absorption effects, but does not
                            consider whether or not the ray was vignetted.
              Exr,Eyr,Ezr : real parts of the electric field components
              Exi,Eyi,Ezi : imaginary parts of the electric field components

          For unpolarized rays --
              error       : (see above)
              intensity   : (see above)

        Note:
        1. The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
           values are accepted.
        2. There is an important exception to the above rule -- If Ex, Ey, Phax,
           Phay are all zero, Zemax will trace two orthogonal rays and the resul-
           ting transmitted intensity will be averaged.
        3. Always check to verify the ray data is valid (check the error) before
           using the rest of the data in the tuple.
        4. Use of zGetPolTraceDirect() has significant overhead as only one ray
           per DDE call is traced. Please refer to the ZEMAX manual for more
           details. Also, if a large number of rays are to be traced, see the
           section "Tracing large number of rays" in the ZEMAX manual.

        See also zGetPolTraceDirect, zGetTrace, zGetTraceDirect
        """
        args0 = "{wN:d},{m:d},".format(wN=waveNum,m=mode)
        args1 = "{sa:d},{sd:d},".format(sa=startSurf,sd=stopSurf)
        args2 = "{x:1.20g},{y:1.20g},{y:1.20g},".format(x=x,y=y,z=z)
        args3 = "{l:1.20g},{m:1.20g},{n:1.20g},".format(l=l,m=m,n=n)
        args4 = "{Ex:1.4f},{Ey:1.4f}".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = "GetPolTraceDirect," + arg0 + args1 + args2 + args3 + args4 + args5
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        rayPolTraceData = tuple([int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)])
        return rayPolTraceData

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

    def zGetSag(self,surfaceNumber,x,y):
        """Returns the sag of the surface with the number `surfaceNumber`, at
        `x` and `y` coordinates on the surface. The returned `sag` and the
        coordinates `x` and `y` are in lens units.

        zGetSag(surfaceNumber,x,y)->sagTuple

        args:
          surfaceNumber : (integer) surface number
          x             : (float) x coordinate in lens units
          y             : (float) y coordinate in lens units
        ret: sagTuple is a 2-tuple containing the following elements
          sag           : (float) sag of the surface at (x,y) in lens units
          alternateSag  : (float) altenate sag
        """
        cmd = "GetSag,{:d},{:1.20g},{:1.20g}".format(surfaceNumber,x,y)
        reply = self.conversation.Request(cmd)
        sagData = reply.rsplit(",")
        return (float(sagData[0]),float(sagData[1]))

    def zGetSequence(self):
        """Returns the sequence number of the lens in the Server's memory, and
        the sequence number of the lens in the LDE in a 2-tuple.

        zGetSequence()->sequenceNumbers

        args:
          None
        ret:
          sequenceNumbers : 2-tuple containing the sequence numbers
        """
        reply = self.conversation.Request("GetSequence")
        seqNum = reply.rsplit(",")
        return (float(seqNum[0]),float(seqNum[1]))

    def zGetSerial(self):
        """Get the serial number
        """
        reply = self.conversation.Request('GetSerial')
        return int(reply.rstrip())

    def zGetSettingsData(self,tempFile,number):
        """Returns the settings data used by a window.

        The data must have been previously stored by a call to zSetSettingsData()
        or the data may have been stored by a previous execution of the client
        program.

        zGetSettingsData(tempFile,number)->settingsData

        args:
          tempfile  : the name of the output file passed by ZEMAX to the client.
                      ZEMAX uses this name to identify for which window the
                      zGetSettingsData() request is for.
          number    : the data number used by the previous zSetSettingsData call.
                      Currently, only number = 0 is supported.
        ret:
          settingsData : (string)  the string that was saved by a previous
                         zSetSettingsData() function for the window & number.

        See also zSetSettingsData
        """
        cmd = "GetSettingsData,{},{:d}".format(tempFile,number)
        reply = self.conversation.Request(cmd)
        return str(reply.rstrip())

    def zGetSolve(self,surfaceNumber,code):
        """Returns data about solves and/or pickups on the surface with number
        `surfaceNumber`.

        zGetSolve(surfaceNumber,code)->solveData

        args:
          surfaceNumber : (integer) surface number
          code          : (integer) indicating which surface parameter the solve
                           is for, such as curvature, thickness, glass, conic,
                           semi-diameter, etc. (see the table below)
        ret:
          solveData     : a tuple, depending on the code value according to the
                          following table if successful, else -1.
        ------------------------------------------------------------------------
        GetSolve Code           -  Returned data format
        ------------------------------------------------------------------------
        0 (curvature)           -  solvetype, parameter1, parameter2, pickupcolumn
        1 (thickness)           -  solvetype, parameter1, parameter2, parameter3,
                                   pickupcolumn
        2 (glass)               -  solvetype (for solvetype = 0)
                                   solvetype, Index, Abbe, Dpgf (for solvetype = 1,
                                   model glass)
                                   solvetype, pickupsurf (for solvetype = 2, pickup)
                                   solvetype, index_offset, abbe_offset (for
                                   solvetype = 4, offset)
                                   solvetype (for solvetype = all other values)
        3 (semi-diameter)        - solvetype, pickupsurf, pickupcolumn
        4 (conic)                - solvetype, pickupsurf, pickupcolumn
        5-16 (parameters 1-12)   - solvetype, pickupsurf, offset,  scalefactor,
                                   pickupcolumn
        17 (parameter 0)         - solvetype, pickupsurf, offset,  scalefactor,
                                   pickupcolumn
        1001+ (extra data values 1+) - solvetype, pickupsurf, scalefactor, offset,
                                       pickupcolumn
        ------------------------------------------------------------------------
        Note: The `solvetype` is an integer code, & the parameters have meanings
        that depend upon the solve type; see the chapter "SOLVES" in the Zemax
        manual for details.

        See also zSetSolve, zGetNSCSolve, zSetNSCSolve.
        """
        cmd = "GetSolve,{:d},{:d}".format(surfaceNumber,code)
        reply = self.conversation.Request(cmd)
        solveData = process_get_set_Solve(reply)
        return solveData

    def zGetSurfaceData(self,surfaceNumber,code,arg2=None):
        """Gets surface data on a sequential lens surface.

        zGetSurfaceData(surfaceNum,code [, arg2])-> surfaceDatum

        args:
            surfaceNum : the surface number
            code       : integer number (see below)
            arg2       : (Optional, integer) for item codes above 70.

        Gets surface datum at surfaceNumber depending on the code according to
        the following table.
        The code is as shown in the following table.
        arg2 is required for some item codes. [Codes above 70]
        ------------------------------------------------------------------------
        Code      - Data returned by zGetSurfaceData()
        ------------------------------------------------------------------------
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
        ------------------------------------------------------------------------
        See also zSetSurfaceData, zGetSurfaceParameter and ZemaxSurfTypes
        """
        if arg2== None:
            cmd = "GetSurfaceData,{sN:d},{c:d}".format(sN=surfaceNumber,c=code)
        else:
            cmd = "GetSurfaceData,{sN:d},{c:d},{a:d}".format(sN=surfaceNumber,
                                                                 c=code,a=arg2)
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
        cmd = "GetSurfaceDLL,{sN:d}".format(surfaceNumber)
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
        cmd = "GetSurfaceParameter,{sN:d},{p:d}".format(sN=surfaceNumber,p=parameter)
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
                                     0 = entrance pupil diameter (EPD)
                                     1 = image space F/#         (IF/#)
                                     2 = object space NA         (ONA)
                                     3 = float by stop           (FBS)
                                     4 = paraxial working F/#    (PWF/#)
                                     5 = object cone angle       (OCA)
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
        systemAperData = tuple([float(elem) if i==2 else int(float(elem))
                                for i, elem in enumerate(rs)])
        return systemAperData

    def zGetSystemProperty(self,code):
        """Returns properties of the system, such as system aperture, field,
        wavelength, and other data, based on the integer `code` passed.

        zGetSystemProperty(code)-> sysPropData
        args:
            code        : (integer) value that defines the specific system property
                          requested (see below).
        ret:
            sysPropData : Returned system property data. Either a string or numeric
                          data.

        This function mimics the ZPL function SYPR.
        ------------------------------------------------------------------------
        Code    Property (the values in the bracket are the expected returns)
        ------------------------------------------------------------------------
          4   - Adjust Index Data To Environment. (0:off, 1:on.)
         10   - Aperture Type code. (0:EPD, 1:IF/#, 2:ONA, 3:FBS, 4:PWF/#, 5:OCA)
         11   - Aperture Value. (stop surface semi-diameter if aperture type is FBS,
                else system aperture)
         12   - Apodization Type code. (0:uniform, 1:Gaussian, 2:cosine cubed)
         13   - Apodization Factor.
         14   - Telecentric Object Space. (0:off, 1:on)
         15   - Iterate Solves When Updating. (0:off, 1:on)
         16   - Lens Title.
         17   - Lens Notes.
         18   - Afocal Image Space. (0:off or "focal mode", 1:on or "afocal mode")
         21   - Global coordinate reference surface.
         23   - Glass catalog list. (Use a string or string variable with the glass
                catalog name, such as "SCHOTT". To specify multiple catalogs use
                a single string or string variable containing names separated by
                spaces, such as "SCHOTT HOYA OHARA".)
         24   - System Temperature in degrees Celsius.
         25   - System Pressure in atmospheres.
         26   - Reference OPD method. (0:absolute, 1:infinity, 2:exit pupil, 3:absolute 2.)
         30   - Lens Units code. (0:mm, 1:cm, 2:inches, 3:Meters)
         31   - Source Units Prefix. (0:Femto, 1:Pico, 2:Nano, 3:Micro, 4:Milli,
                5:None,6:Kilo, 7:Mega, 8:Giga, 9:Tera)
         32   - Source Units. (0:Watts, 1:Lumens, 2:Joules)
         33   - Analysis Units Prefix. (0:Femto, 1:Pico, 2:Nano, 3:Micro, 4:Milli,
                5:None,6:Kilo, 7:Mega, 8:Giga, 9:Tera)
         34   - Analysis Units "per" Area. (0:mm^2, 1:cm^2, 2:inches^2, 3:Meters^2, 4:feet^2)
         35   - MTF Units code. (0:cycles per millimeter, 1:cycles per milliradian.
         40   - Coating File name.
         41   - Scatter Profile name.
         42   - ABg Data File name.
         43   - GRADIUM Profile name.
         50   - NSC Maximum Intersections Per Ray.
         51   - NSC Maximum Segments Per Ray.
         52   - NSC Maximum Nested/Touching Objects.
         53   - NSC Minimum Relative Ray Intensity.
         54   - NSC Minimum Absolute Ray Intensity.
         55   - NSC Glue Distance In Lens Units.
         56   - NSC Missed Ray Draw Distance In Lens Units.
         57   - NSC Retrace Source Rays Upon File Open. (0:no, 1:yes)
         58   - NSC Maximum Source File Rays In Memory.
         59   - Simple Ray Splitting. (0:no, 1:yes)
         60   - Polarization Jx.
         61   - Polarization Jy.
         62   - Polarization X-Phase.
         63   - Polarization Y-Phase.
         64   - Convert thin film phase to ray equivalent. (0:no, 1:yes)
         65   - Unpolarized. (0:no, 1:yes)
         66   - Method. (0:X-axis, 1:Y-axis, 2:Z-axis)
         70   - Ray Aiming. (0:off, 1:on, 2:aberrated)
         71   - Ray aiming pupil shift x.
         72   - Ray aiming pupil shift y.
         73   - Ray aiming pupil shift z.
         74   - Use Ray Aiming Cache. (0:no, 1:yes)
         75   - Robust Ray Aiming. (0:no, 1:yes)
         76   - Scale Pupil Shift Factors By Field. (0:no, 1:yes)
         77   - Ray aiming pupil compress x.
         78   - Ray aiming pupil compress y.
         100  - Field type code. (0=angl,1=obj ht,2=parx img ht,3=rel img ht)
         101  - Number of fields.
         102,103 - The field number is value1, value2 is the field x, y coordinate
         104  - The field number is value1, value2 is the field weight
         105,106 - The field number is value1, value2 is the field vignetting
                   decenter x, decenter y
         107,108 - The field number is value1, value2 is the field vignetting
                   compression x, compression y
         109  - The field number is value1, value2 is the field vignetting angle
         110  - The field normalization method, value 1 is 0 for radial and 1 for
                rectangular
         200  - Primary wavelength number.
         201  - Number of wavelengths
         202  - The wavelength number is value1, value 2 is the wavelength in
                micrometers.
         203  - The wavelength number is value1, value 2 is the wavelength weight
         901  - The number of CPU's to use in multi-threaded computations, such as
                optimization. (0=default). See the manual for details.

        NOTE: Currently Zemax returns just "0" for the codes: 102,103, 104,105,
              106,107,108,109, and 110. This is unexpected! So, PyZDDE will return
              the reply (string) as is for the user to handle.
        ------------------------------------------------------------------------
        See also zSetSystemProperty
        """
        cmd = "GetSystemProperty,{c}".format(c=code)
        reply = self.conversation.Request(cmd)
        sysPropData = process_get_set_SystemProperty(code,reply)
        return sysPropData

    def zGetTextFile(self,textFileName, analysisType, settingsFileName, flag):
        """Request to save a text file for any analysis that supports text output.

           zGetText(textFilename, analysisType, settingsFileName, flag) -> retVal

           args:
            textFileName : name of the file to be created including the full path,
                           name, and extension for the text file.
            analysisType : 3 letter case-sensitive label that indicates the
                           type of the analysis to be performed. They are identical
                           to those used for the button bar in Zemax. The labels
                           are case sensitive. If no label is provided or recognized,
                           a standard raytrace will be generated.
            settingsFileName : If a valid file name is used for the "settingsFileName",
                               ZEMAX will use or save the settings used to compute the
                               text file, depending upon the value of the flag parameter.
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
            cmd = 'GetTextFile,"{tF}",{aT},"{sF}",{fl:d}'.format(tF=textFileName,
                                    aT=analysisType,sF=settingsFileName,fl=flag)
            reply = self.conversation.Request(cmd)
            if 'OK' in reply.split():
                retVal = 0
        return retVal

    def zGetTol(self,operandNumber):
        """Returns the tolerance data.

        zGetTol(operandNumber)->toleranceData

        args:
          operandNumber : 0 or the tolerance operand number (row number in the
                          tolerance editor, when greater than 0)
        ret:
          toleranceData. It is a number or a 6-tuple, depending upon `operandNumber`
          as follows:
            if operandNumber == 0, toleranceData = number where `number` is the
            number of tolerance operands defined.
            if operandNumber > 0, toleranceData = (tolType, int1, int2, min, max, int3)

        See also zSetTol, zSetTolRow
        """
        reply = self.conversation.Request("GetTol,{:d}".format(operandNumber))
        if operandNumber == 0:
            toleranceData = int(float(reply.rstrip()))
            if toleranceData == 1:
                reply = self.conversation.Request("GetTol,1")
                tolType = reply.rsplit(",")[0]
                if tolType == 'TOFF': # the tol editor is actually empty
                    toleranceData = 0
        else:
            toleranceData = process_get_set_Tol(operandNumber,reply)
        return toleranceData

    def zGetTrace(self,waveNum,mode,surf,hx,hy,px,py):
        """Trace a (single) ray through the current lens in the ZEMAX DDE server.

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

        Example: To trace the real chief ray to surface 5 for wavelength 3, use
            rayTraceData = zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

            OR

            (errorCode,vigCode,x,y,z,l,m,n,l2,m2,n2,intensity) = \
                                              zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

        Note:
        1. Always check to verify the ray data is valid  (errorCode) before using
           the rest of the string!
        2. Use of zGetTrace() has significant overhead as only one ray per DDE call
           is traced. Please refer to the ZEMAX manual for more details. Also, if a
           large number of rays are to be traced, see the section "Tracing large
           number of rays" in the ZEMAX manual.

        See also zGetTraceDirect, zGetPolTrace, zGetPolTraceDirect
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f}".format(px=px,py=py)
        cmd = "GetTrace," + args1 + args2 + args3
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        rayTraceData = tuple([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return rayTraceData

    def zGetTraceDirect(self,waveNum,mode,startSurf,stopSurf,x,y,z,l,m,n):
        """Trace a (single) ray through the current lens in the ZEMAX DDE server
        while providing a more direct access to the ZEMAX ray tracing engine than
        zGetTrace.

        zGetTraceDirect(waveNum,mode,startSurf,stopSurf,x,y,z,l,m,n)->rayTraceData

        args:
          waveNum  : wavelength number as in the wavelength data editor
          mode     :  0 = real, 1 = paraxial
          startSurf: starting surface of the ray
          stopSurf : stopping surface of the ray
          x,y,z,   : coordinates of the ray at the starting surface
          l,m,n    : the direction cosines to the entrance pupil aim point for
                     the x-, y-, z- direction cosines respectively
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
              intensity : the relative transmitted intensity of the ray, excluding
                          any pupil apodization. The surface apodization is defined.
        Notes:
         Normally, rays are defined by the normalized field and pupil coordinates
         hx, hy, px, and py. ZEMAX takes these normalized coordinates and computes
         the object coordinates (x, y, and z) and the direction cosines to the
         entrance pupil aim point (l, m, and n; for the x-, y-, and z-direction
         cosines, respectively).However, there are times when it is more appropriate
         to trace rays by direct specification of x, y, z, l, m, and n. The direct
         specification has the added flexibility of defining the starting surface
         for the ray anywhere in the optical system.
        """
        args1 = "{wN:d},{m:d},".format(wN=waveNum,m=mode)
        args2 = "{sa:d},{sp:d},".format(sa=startSurf,sp=stopSurf)
        args3 = "{x:1.20f},{y:1.20f},{z:1.20f}".format(x=x,y=y,z=z)
        args4 = "{l:1.20f},{m:1.20f},{n:1.20f}".format(l=l,m=m,n=n)
        cmd = "GetTraceDirect," + args1 + args2 + args3 + args4
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        rayTraceData = tuple([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return rayTraceData

    def zGetUDOSystem(self,bufferCode):
        """Load a particular lens from the optimization function memory into the
        ZEMAX server's memory. This will cause ZEMAX to retrieve the correct lens
        from system memory, and all subsequent DDE calls will be for actions
        (such as ray tracing) on this lens.The only time this item name should be
        used is when implementing a User Defined Operand, or UDO.

        zGetUDOSystem(bufferCode)->

        args:
          bufferCode: (integer) The buffercode is an integer value provided by
                      ZEMAX to the client that uniquely identifies the correct lens.

        ret:
          ?

        Note: Once the data is computed, up to 1001 values may be sent back to
        the server, and ultimately to the optimizer within ZEMAX, with the
        zSetUDOItem command.

        See also zSetUDOItem.
        """
        cmd = "GetUDOSystem,{:d}".format(bufferCode)
        reply = self.conversation.Request(cmd)
        return regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zGetUpdate(self):
        """Update the lens, which means Zemax recomputes all pupil positions,
        solves, and index data.

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
            cmd = "GetWave,{wC:d}".format(wC=i+1)
            reply = self.conversation.Request(cmd)
            rs = reply.split(',')
            waveDataTuple[0].append(float(rs[0])) # store the wavelength
            waveDataTuple[1].append(float(rs[1])) # store the weight
        return (tuple(waveDataTuple[0]),tuple(waveDataTuple[1]))

    def zHammer(self,numOfCycles,algorithm):
        """Calls the Hammer optimizer.

        zHammer(numOfCycles,algorithm)->finalMeritFn

        args:
          numOfCycles  : (integer) the number of cycles to run
                         if numOfCycles < 1, zHammer updates all operands in the
                         merit function and returns the current merit function,
                         and no optimization is performed.
          algorithm    : 0 = Damped Least Squares
                         1 = Orthogonal descent
        ret:
          finalMeritFn : (float) the final merit function.

        Note:
          1. If the merit function value returned is 9.0E+009, the optimization
             failed, usually because the lens or merit function could not be
             evaluated.
          2. The number of cycles should be kept small enough to allow the algorithm
             to complete and return before the DDE communication times out, or
             an error will occur. One possible way to achieve high number of
             cycles could be to call zHammer multiple times in a loop, each time
             comparing the returned merit function with few of the previously
             returned (& stored) merit functions to determine if an optimum has
             been attained.

        See also zOptimize,  zLoadMerit, zsaveMerit
        """
        cmd = "Hammer,{:1.2g},{:d}".format(numOfCycles,algorithm)
        reply = self.conversation.Request(cmd)
        return float(reply.rstrip())

    def zImportExtraData(self,surfaceNumber,fileName):
        """Imports extra data and grid surface data values into an existing sur-
        face.

        zImportExtraData(surfaceNumber,fileName)->

        args:
          surfaceNumber : (integer) surface number
          fileName      : (string) file name (of an ASCII file)
        ret:

        Note:
            The ASCII file is a single column of free-format numbers, with a
            .DAT extension.
        """
        cmd = "ImportExtraData,{:d},{}".format(surfaceNumber,fileName)
        reply = self.conversation.Request(cmd)
        return reply.rstrip()
        # !!! FIX determine what is the currect return


    def zInsertConfig(self,configNumber):
        """Insert a new configuration (column) in the multi-configuration editor.
        The new configuration will be placed at the location (column) indicated
        by the parameter `configNumber`.

        zInsertConfig(configNumber)->configNumberRet

        args:
            configNumber    : (integer) the configuration (column) number to insert.
        ret:
            configNumberRet : (integer) the column number of the configuration that
                              is inserted at configNumber.

        Note:
            1. The configNumber returned (configNumberRet) is generally different
               from the number in the input configNumber.
            2. Use zInsertMCO() to insert a new multi-configuration operand in the
               multi-configuration editor.
            3. Use zSetConfig() to switch the current configuration number

        See also zDeleteConfig.
        """
        return int(self.conversation.Request("InsertConfig,{:d}".format(configNumber)))

    def zInsertMCO(self,operandNumber):
        """Insert a new multi-configuration operand (row) in the multi-configuration
        editor.

        zInsertMCO(operandNumber)-> retValue

        args:
            operandNumber : (integer) between 1 and the current number of operands
                            plus 1, inclusive.
        ret:
            retValue      : new number of operands (rows).

        See also zDeleteMCO. Use zInsertConfig(), to insert a new configuration (row).
        """
        return int(self.conversation.Request("InsertMCO,{:d}".format(operandNumber)))

    def zInsertMFO(self,operandNumber):
        """Insert a new optimization operand (row) in the merit function editor.

        zInsertMFO(operandNumber)->retValue

        args:
            operandNumber : (integer) between 1 and the current number of operands
                            plus 1, inclusive.
        ret:
            retValue      : new number of operands (rows).

        See also zDeleteMFO. Generally, you may want to use zSetOperand() afterwards.
        """
        return int(self.conversation.Request("InsertMFO,{:d}".format(operandNumber)))

    def zInsertObject(self,surfaceNumber,objectNumber):
        """
        Insert a new NSC object at the location indicated by the parameters
        `surfaceNumber` and `objectNumber`.

        zInsertObject(surfaceNumber,objectNumber)->status

        args:
          surfaceNumber : (integer) surface number of the NSC group. Use 1 if
                          the program mode is Non-Sequential.
          objectNumber  : object number
        ret:
          status        : 0 if successful, -1 if failed.

        See also zSetNSCObjectData to define data for the new surface and the
        zDeleteObject function.
        """
        cmd = "InsertObject,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self.conversation.Request(cmd)
        if reply.rstrip() == 'BAD COMMAND':
            return -1
        else:
            return int(reply.rstrip())

    def zInsertSurface(self,surfNum):
        """Insert a lens surface in the ZEMAX DDE server. The new surface will be
        placed at the location indicated by the parameter surfNum.

        zInsertSruface(surfNum)-> retVal (0 = success)

        See also zSetSurfaceData() to define data for the new surface and the
        zDeleteSurface() functions.
        """
        return int(self.conversation.Request("InsertSurface,"+str(surfNum)))

    def zLoadDetector(self, surfaceNumber, objectNumber, fileName):
        """Loads the data saved in a file to an NSC Detector Rectangle, Detector
        Color, Detector Polar, or Detector Volume object.

        zLoadDetector(surfaceNumber, objectNumber, fileName)->status

        args:
          surfNumber   : (integer) surface number of the non-sequential group.
                         Use 1 if the program mode is Non-Sequential.
          objectNumber : (integer) object number of the detector object
          fileName     : (string) The filename may include the full path, if no
                         path is provided the path of the current lens file is
                         used. The extension should be DDR, DDC, DDP, or DDV for
                         Detector Rectangle, Color, Polar, and Volume objects,
                         respectively.
        ret:
          status       : 0 if load was successful
                         Error code (such as -1,-2) if failed.
        """
        isRightExt = path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        isFile = path.isfile(fileName)  # check if file exist
        if isRightExt and isFile:
            cmd = ("LoadDetector,{:d},{:d},{}"
                   .format(surfaceNumber,objectNumber,fileName))
            reply = self.conversation.Request(cmd)
            return regressLiteralType(reply.rstrip())
        else:
            return -1

    def zLoadFile(self,fileName,append=None):
        """Loads a ZEMAX file into the server.

        zLoadFile(fileName[,append]) -> retVal

        args:
            filename: full path of the ZEMAX file to be loaded. For example:
                      "C:\ZEMAX\Samples\cooke.zmx"
            append  : (optional, integer) If a non-zero value of append is passed,
                      then the new file is appended to the current file starting
                      at the surface number defined by the value appended.
        retVal:
                0: file successfully loaded
             -999: file could not be loaded (check if the file really exists, or
                   check the path.
             -998: the command timed out
            other: the upload failed.

        See also zSaveFile, zGetPath, zPushLens, zuiLoadFile
        """
        reply = None
        isAbsPath = path.isabs(fileName)
        isRightExt = path.splitext(fileName)[1] in ('.zmx','.ZMX')
        isFile = path.isfile(fileName)
        if isAbsPath and isRightExt and isFile:
            if append:
                cmd = "LoadFile,{},{}".format(fileName,append)
            else:
                cmd = "LoadFile,{}".format(fileName)
            reply = self.conversation.Request(cmd)
            if reply:
                return int(reply) #Note: Zemax returns -999 if update fails.
            else:
                return -998
        else:
            return -999

    def zLoadMerit(self, fileName):
        """Loads a ZEMAX .MF or .ZMX file and extracts the merit function and
        places it in the lens loaded in the server.

        zLoadMerit(fileName)->meritData

        args:
          fileName  : (string) name of the merit function file with full path and
                       extension.
        ret:
          meritData : If the loading is successful, meritData is a 2-tuple
                      containing the following elements:
                      number : (integer) number of operands in the merit function
                      merit  : (float) merit value of the merit function.

                      If meritData = -999, file could not be loaded (check if the
                      file really exists, or check the path.
        Note:
          1. If the merit function value is 9.00e+009, the merit function cannot
             be evaluated.
          2. Loading a merit function file does not change the data displayed in
             the LDE; the server process has a separate copy of the lens data.

        See also: zOptimize, zSaveMerit
        """
        isAbsPath = path.isabs(fileName)
        isRightExt = path.splitext(fileName)[1] in ('.mf','.MF','.zmx','.ZMX')
        isFile = path.isfile(fileName)
        if isAbsPath and isRightExt and isFile:
            reply = self.conversation.Request('LoadMerit,'+fileName)
            rs = reply.rsplit(",")
            meritData = [int(float(e)) if i==0 else float(e)
                         for i,e in enumerate(rs)]
            return tuple(meritData)
        else:
            return -999

    def zLoadTolerance(self, fileName):
        """Loads a tolerance file previously saved with zSaveTolerance and
        places the tolerances in the lens loaded in the DDE server.

        zLoadTolerance(fileName)->numTolOperands

        args:
          fileName : (string) file name of the tolerance file. If no path is
                     provided in the filename, the <data>\Tolerance folder is
                     assumed.
        ret:
          numTolOperands : number of tolerance operands loaded.
                           -999 if file doesnot exist
        """
        if path.isabs(fileName): # full path is provided
            fullFilePathName = fileName
        else:                    # full path not provided
            fullFilePathName = self.zGetPath()[0] + "\\Tolerance\\" + fileName
        if path.isfile(fullFilePathName):
            cmd = "LoadTolerance,{}".format(fileName)
            reply = self.conversation.Request(cmd)
            return int(float(reply.rstrip()))
        else:
            return -999

    def zMakeGraphicWindow(self,fileName,moduleName,winTitle,textFlag,settingsData=None):
        """Notifies ZEMAX that graphic data has been written to a file and may now
        be displayed as a ZEMAX child window. The primary purpose of this item is
        to implement user defined features in a client application, that look and
        act like native ZEMAX features.

        There are two ways of using this command:

        zMakeGraphicWindow(fileName,moduleName,winTitle,textFlag,settingsData)->retString

          OR

        zSetSettingsData(0,settingsData)
        zMakeGraphicWindow(self,fileName,moduleName,winTitle,textFlag)->retString

        args:
          fileName     : the full path and file name to the temporary file that
                         holds the graphic data. This must be the same name as
                         passed to the client executable in the command line
                         arguments, if any.
          moduleName   : the full path and executable name of the client program
                         that created the graphic data.
          winTitle     : the string which defines the title ZEMAX should place in
                         the top bar of the window.
          textFlag     : 1 => the client can also generate a text version of the
                         data. Since the current data is a graphic display (it
                         must be if the item is MakeGraphicWindow) ZEMAX wants
                         to know if the "Text" menu option should be available
                         to the user, or if it should be grayed out.
                         0 => ZEMAX will gray out the "Text" menu option and will
                         not attempt to ask the client to generate a text version
                         of the data.
          settingsData : The settings data is a string of values delimited by
                         spaces (not commas) which are used by the client to define
                         how the data was generated. These values are only used
                         by the client, not by ZEMAX. The settings data string
                         holds the options and data that would normally appear in
                         a ZEMAX "settings" style dialog box. The settings data
                         should be used to recreate the data if required. Because
                         the total length of a data item cannot exceed 255
                         characters, the function zSetSettingsData() may be used
                         prior to the call to zMakeGraphicWindow() to specify the
                         settings data string rather than including the data as
                         part of zMakeGraphicWindow(). See "How ZEMAX calls the
                         client" in the manual for more details on the settings
                         data.

        A sample item string might look like the following:

        zMakeGraphicWindow('C:\TEMP\ZGF001.TMP','C:\ZEMAX\FEATURES\CLIENT.EXE',
                           'ClientWindow',1,"0 1 2 12.55")

        This call indicates that ZEMAX should open a graphic window, display the
        data stored in the file 'C:\TEMP\ZGF001.TMP', and that any updates or
        setting changes can be made by calling the client module
        'C:\ZEMAX\FEATURES\CLIENT.EXE'. This client can generate a text version
        of the graphic, and the settings data string (used only by the client)
        is "0 1 2 12.55".
        """
        if settingsData:
            cmd = ("MakeGraphicWindow,{},{},{},{:d},{}"
                   .format(fileName,moduleName,winTitle,textFlag,settingsData))
        else:
            cmd = ("MakeGraphicWindow,{},{},{},{:d}"
                   .format(fileName,moduleName,winTitle,textFlag))
        reply = self.conversation.Request(cmd)
        return str(reply.rstrip())
        # FIX !!! What is the appropriate reply?

    def zMakeTextWindow(self,fileName,moduleName,winTitle,settingsData=None):
        """Notifies ZEMAX that text data has been written to a file and may now
        be displayed as a ZEMAX child window. The primary purpose of this item
        is to implement user defined features in a client application, that look
        and act like native ZEMAX features.

        There are two ways of using this command:

        zMakeTextWindow(fileName,moduleName,winTitle,settingsData)->retString

          OR

        zSetSettingsData(0,settingsData)
        zMakeTextWindow(self,fileName,moduleName,winTitle)->retString

        args:
          fileName     : the full path and file name to the temporary file that
                         holds the text data. This must be the same name as
                         passed to the client executable in the command line
                         arguments, if any.
          moduleName   : the full path and executable name of the client program
                         that created the text data.
          winTitle     : the string which defines the title ZEMAX should place in
                         the top bar of the window.
          settingsData : The settings data is a string of values delimited by
                         spaces (not commas) which are used by the client to define
                         how the data was generated. These values are only used
                         by the client, not by ZEMAX. The settings data string
                         holds the options and data that would normally appear in
                         a ZEMAX "settings" style dialog box. The settings data
                         should be used to recreate the data if required. Because
                         the total length of a data item cannot exceed 255
                         characters, the function zSetSettingsData() may be used
                         prior to the call to zMakeTextWindow() to specify the
                         settings data string rather than including the data as
                         part of zMakeTextWindow(). See "How ZEMAX calls the
                         client" in the manual for more details on the settings
                         data.

        A sample item string might look like the following:

        zMakeTextWindow('C:\TEMP\ZGF002.TMP','C:\ZEMAX\FEATURES\CLIENT.EXE',
                           'ClientWindow',"6 5 4 12.55")

        This call indicates that ZEMAX should open a text window, display the
        data stored in the file 'C:\TEMP\ZGF002.TMP', and that any updates or
        setting changes can be made by calling the client module
        'C:\ZEMAX\FEATURES\CLIENT.EXE'. The settingsdata string (used only by the
        client) is "6 5 4 12.55".
        """
        if settingsData:
            cmd = ("MakeTextWindow,{},{},{},{}"
                   .format(fileName,moduleName,winTitle,settingsData))
        else:
            cmd = ("MakeTextWindow,{},{},{}"
                   .format(fileName,moduleName,winTitle))
        reply = self.conversation.Request(cmd)
        return str(reply.rstrip())
        # FIX !!! What is the appropriate reply?

    def zModifySettings(self,fileName,mType,value):
        """Used to change specific options in ZEMAX settings files.

        The settings files are used by zMakeTextWindow() and zMakeGraphicWindow()

        zModifySettings(fileName,mType,value)->status

        args:
          fileName : The full name of the settings file, including the path.
          mType    : a mnemonic that defines what option value is being modified.
                     The valid values for type are as defined in the ZPL macro
                     command MODIFY-SETTINGS that serves the same function as
                     zModifySettings() does for extensions. See "MODIFYSETTINGS"
                     in the Zemax manual for a complete list of the type codes.
          value    : value
        ret:
          status :  0 = no error
                   -1 = invalid file
                   -2 = incorrect version number
                   -3 = file access conflict
        """
        cmd = "ModifySettings,{},{},{:1.20g}".format(fileName,mType,value)
        reply = self.conversation.Request(cmd)
        return int(float(reply.rstrip()))

    def zNewLens(self):
        """Erases the current lens.

        The "minimum" lens that remains is identical to the lens Data Editor
        when "File,New" is selected. No prompt to save the existing lens is given.

        zNewLens-> retVal (retVal = 0 means successful)
        """
        return int(self.conversation.Request('NewLens'))

    def zNSCCoherentData(self,surfaceNumber,objectNumDisDetectr,pixel,dataType):
        """Return data from an NSC detector (Non-sequential coherent data)

        zNSCCoherentData(surfaceNumber,objectNumDisDetectr,pixel,dataType)->

        args:
          surfaceNumber       : The surface number of the NSC group (1 for pure
                                NSC systems).
          objectNumDisDetectr : The object number of the desired detector.
          pixel               : 0 = the sum of the data for all pixels for that
                                    detector object is returned
                                +ve int = the data from the specified pixel is
                                          returned.
          dataType            : 0 = real
                                1 = imaginary
                                2 = amplitude
                                3 = power
        """
        cmd = ("NSCCoherentData,{:d},{:d},{:d},{:d}"
               .format(surfaceNumber,objectNumDisDetectr,pixel,dataType))
        reply = self.conversation.Request(cmd)
        return float(reply.rstrip())

    def zNSCDetectorData(self,surfaceNumber,objectNumDisDetectr,pixel,dataType):
        """Return data from an NSC detector (Non-sequential incoherent intensity
        data)

        zNSCDetectorData(surfaceNumber,objectNumDisDetectr,pixel,dataType)->

        args:
          surfaceNumber       : The surface number of the NSC group (1 for pure
                                NSC systems).
          objectNumDisDetectr : The object number of the desired detector.
                                0 = all detectors are cleared
                                -ve int = only the detector defined by the absolute
                                          value of `objectNumDisDetectr` is cleared

          For Detector Rectangles, Detector Surfaces, & all faceted detectors:
          -------------------------------------------------------------------
          pixel : +ve int = the data from the specified pixel is returned.
                        0 = the sum of the total flux in position space,average flux/area
                            in position space, or total flux in angle space for all pixels
                            for that detector object, for Data = 0, 1, or 2, respectively.
                       -1 = Maximum flux or flux/area.
                       -2 = Minimum flux or flux/area.
                       -3 = Number of rays striking the detector.
                       -4 = Standard deviation (RMS from the mean) of all the non-zero pixel data.
                       -5 = The mean value of all the non-zero pixel data.
                       -6,-7,-8 = The x, y, or z coordinate of the position or angle Irradiance
                                  or Intensity  centroid, respectively.
                       -9,-10,-11,-12,-13 = The RMS radius, x, y, z, or xy cross term distance
                                            or angle of all the pixel data with respect to the
                                            centroid. These are the second moments r^2, x^2, y^2,
                                            z^2, & xy, respectively.
          dataType : 0 = flux
                     1 = flux/area
                     2 = flux/solid angle pixel
                     Note:
                       Only values 0 & 1 (for flux & flux/area) are supported for faceted detectors.

          For Detector Volumes:
          -----------------------
          pixel : is interpreted as the voxel number.
                  if pixel == 0,  the value returned is the sum for all pixels.
          dataType: 0 = incident flux
                    1 = absorbed flux
                    2 = absorbed flux per unit volume
        """
        cmd = ("NSCDetectorData,{:d},{:d},{:d},{:d}"
               .format(surfaceNumber,objectNumDisDetectr,pixel,dataType))
        reply = self.conversation.Request(cmd)
        return float(reply.rstrip())

    def zNSCTrace(self,surfNum,objNumSrc,split=0,scatter=0,usePolar=0,
                  ignoreErrors=0,randomSeed=0,save=0,saveFilename=None,oFilter=None):
        """Traces rays from one or all NSC sources with various optional arguments.
        zNSCTrace() always updates the lens before tracing rays to make certain all
        objects are correctly loaded and updated.

        Possible ways of using this function:

        zNSCTrace(surfNum,objNumSrc,[split,scatter,usePolar,ignoreErrors,
                  randomSeed,save=0])->traceResult

          OR

        zNSCTrace(surfNum,objNumSrc,[split,scatter,usePolar,ignoreErrors,
                  randomSeed,]save,saveFilename,oFilter)->traceResult

        args:
          surfNum      : The surface number of the NSC group (1 for pure NSC systems).
          objNumSrc    : The object number of the desired source.
                         0 = all sources will be traced.
          split        : 0 = splitting is OFF (default)
                        otherwise = splitting is ON
          scatter      : 0 = scattering is OFF (default)
                        otherwise = scattering is ON
          usePolar     : 0 = polarization is OFF (default)
                        otherwise = polarization is
                        Note: If splitting is ON polarization is automatically selected.
          ignoreErrors : 0 = ray errors will terminate the NSC trace & macro execution
                             and an error will be reported (default).
                        otherwise = erros will be ignored
          randomSeed   : 0 or omitted = the random number generator will be seeded
                            with a random value, & every call to zNSCTrace will
                            produce different random rays (default).
                         integer other than zero = then the random number generator
                            will be seeded with the specified value, and every call
                            to zNSCTrace using the same seed will produce identical rays.
          save         : 0 or omitted = the parameters `saveFilename` and `oFilter`
                            need not be supplied (default).
                         otherwise = the rays will be saved in a ZRD file. The ZRD file
                         will have the name specified by the `saveFilename`, and will
                         be placed in the same directory as the lens file. The extension
                         of `saveFilename` should be ZRD, and no path should be specified.
          saveFilename : (see above)
          oFilter      : If save is not zero,then the optional filter name is either a
                         string variable with the filter, or the literal filter in
                         double quotes. For information on filter strings see
                         "The filter string" in the Zemax manual.
        ret:
          traceResult  : 0 if successful, -1 if problem with saveFileName, other
                         error codes sent by Zemax.

        Examples (the first two examples are from MZDDE):
        -----------------------------------------------
        zNSCTrace(1, 2, 1, 0, 1, 1)

        The above command traces rays in NSC group 1, from source 2, with ray splitting,
        no ray scattering, using polarization and ignoring errors.

        zNSCTrace(1, 2, 1, 0, 1, 1, 33, 1, "myrays.ZRD", "h2")

        Same as above, only a random seed of 33 is given and the data is saved to the
        file "myrays.ZRD" after filtering as per h2.

        zNSCTrace(1, 2)

        The above command traces rays in NSC group 1, from source 2, witout ray splitting,
        no ray scattering, without using polarization and will not ignore errors.

        """
        requiredArgs = ("{:d},{:d},{:d},{:d},{:d},{:d},{:d},{:d}"
        .format(surfNum,objNumSrc,split,scatter,usePolar,ignoreErrors,randomSeed,save))
        if save:
            isAbsPath = path.isabs(saveFilename)
            isRightExt = path.splitext(saveFilename)[1] in ('.ZRD',)
            if isRightExt and not isAbsPath:
                if oFilter:
                    optionalArgs = ",{},{}".format(saveFilename,oFilter)
                else:
                    optionalArgs = ",{}".format(saveFilename)
                cmd = "NSCTrace,"+requiredArgs+optionalArgs
            else:
                return -1 # either full path present in saveFileName or extension is not .ZRD
        else:
            cmd = "NSCTrace,"+requiredArgs
        reply = self.conversation.Request(cmd)
        if 'OK' in reply.split():
            return 0
        else:
            return int(float(reply.rstrip()))  # return the error code sent by zemax.

    def zOpenWindow(self, analysisType, zplMacro=False):
        """Open a new analysis window on the main ZEMAX screen.

        zOpenWindow(analysisType)->status

        args:
          analysisType  : (string) 3 character case-sensitive label that indicates
                          the type of analysis to be performed. The 3 letter
                          labels are identical to those used for the button bar
                          in ZEMAX. You can see a list of the button codes by
                          importing zemaxbuttons module and calling the showZButtons
                          function in an interactive shell, for example:
                              import zemaxbuttons as zb
                              zb.showZButtonList()
          zplMacro      : (bool) True if the analysisTyppe code is the first 3-letters
                          of a ZPL macro name, else False (default).
                          Note: If the analysiType is ZPL macro, please make sure
                          that the macro exist in the <data>/Macros folder. It is
                          recommended to use zExecuteZPLMacro() if you are trying
                          to execute a ZPL macro.
        ret:
          status  : 0 = successful, -1 = incorrect analysis code
        See also zGetMetaFile, zExecuteZPLMacro
        """
        if zb.isZButtonCode(analysisType) ^ zplMacro:  # XOR operation
            reply = self.conversation.Request("OpenWindow,{}".format(analysisType))
            if 'OK' in reply.split():
                return 0
            else:
                return int(float(reply.rstrip()))  # error code from Zemax
        else:
            return -1 # Incorrect analysisType code

    def zOperandValue(self,operandType,*values):
        """Returns the value of any optimization operand, even if the operand is
        not currently in the merit function.

        zOperandValue(operandType,*values)->operandValue

        args:
          operandType  : a valid optimization operand
          *values      : variable length argument. Possible arguments include
                          int1, int2, data1, data2, data3, data4, data5, data6
        ret:
          operandValue : (float) the value
        """
        if zo.isZOperand(operandType,1) and (0 < len(values) < 9):
            valList = [str(int(elem)) if i in (0,1) else str(float(elem))
                       for i,elem in enumerate(values)]
            arguments = ",".join(valList)
            cmd = "OperandValue," + operandType + "," + arguments
            reply = self.conversation.Request(cmd)
            return float(reply.rstrip())
        else:
            return -1

    def zOptimize(self,numOfCycles,algorithm):
        """Calls the Zemax Damped Least Squares (DLS) optimizer.

        zOptimize(numOfCycles,algorithm)->finalMeritFn

        args:
          numOfCycles  : (integer) the number of cycles to run
                         if numOfCycles == 0, optimization runs in automatic mode.
                         if numOfCycles < 0, zOptimize updates all operands in the
                         merit function and returns the current merit function,
                         and no optimization is performed.
          algorithm    : 0 = Damped Least Squares
                         1 = Orthogonal descent
        ret:
          finalMeritFn : (float) the final merit function.

        Note:
          1. If the merit function value returned is 9.0E+009, the optimization
             failed, usually because the lens or merit function could not be
             evaluated.
          2. The number of cycles should be kept small enough to allow the algorithm
             to complete and return before the DDE communication times out, or
             an error will occur. One possible way to achieve high number of
             cycles could be to call zOptimize multiple times in a loop, each time
             comparing the returned merit function with few of the previously
             returned (& stored) merit functions to determine if an optimum has
             been attained.

        See also zHammer, zLoadMerit, zsaveMerit
        """
        cmd = "Optimize,{:1.2g},{:d}".format(numOfCycles,algorithm)
        reply = self.conversation.Request(cmd)
        return float(reply.rstrip())

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
             -999: the lens could not be pushed into the LDE. (check zPushLensPermission)
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

    def zQuickFocus(self,mode=0,centroid=0):
        """Performs a quick best focus adjustment for the optical system by adjusting
        the back focal distance for best focus. The "best" focus is chosen as a wave-
        length weighted average over all fields. It adjusts the thickness of the
        surface prior to the image surface.

        zQuickFocus([mode,centroid]) -> retVal

        arg:
            mode:
                0: RMS spot size (default)
                1: spot x
                2: spot y
                3: wavefront OPD
            centroid: to specify RMS reference
                0: RMS referenced to the chief ray (default)
                1: RMS referenced to image centroid

        ret:
            retVal: 0 for success.
        """
        retVal = -1
        cmd = "QuickFocus,{mode:d},{cent:d}".format(mode=mode,cent=centroid)
        reply = self.conversation.Request(cmd)
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zReleaseWindow(self, tempFileName):
        """Release locked window/menu mar.

        zReleaseWindow(tempFileName)->status

        args:
          tempFileName  : (string)
        ret:
          status  : 0 = no window is using the filename
                    1 = the file is being used.

        When ZEMAX calls the client to update or change the settings used by
        the client function, the menu bar is grayed out on the window to prevent
        multiple updates or setting changes from being requested simultaneously.
        Normally, when the client code calls the functions zMakeTextWindow() or
        zMakeGraphicWindow(), the menu bar is once again activated. However, if
        during an update or setting change, the new data cannot be computed, then
        the window must be released. The zReleaseWindow() function serves just
        this one purpose. If the user selects "Cancel" when changing the settings,
        the client code should send a zReleaseWindow() call to release the lock
        out of the menu bar. If this command is not sent, the window cannot be
        closed, which will prevent ZEMAX from terminating normally.
        """
        reply = self.conversation.Request("ReleaseWindow,{}".format(tempFileName))
        return int(float(reply.rstrip()))

    def zSaveDetector(self, surfaceNumber, objectNumber, fileName):
        """Saves the data currently on an NSC Detector Rectangle, Detector Color,
        Detector Polar, or Detector Volume object to a file.

        zSaveDetector(surfaceNumber, objectNumber, fileName)->status

        args:
          surfNumber   : (integer) surface number of the non-sequential group.
                         Use 1 if the program mode is Non-Sequential.
          objectNumber : (integer) object number of the detector object
          fileName     : (string) The filename may include the full path, if no
                         path is provided the path of the current lens file is
                         used. The extension should be DDR, DDC, DDP, or DDV for
                         Detector Rectangle, Color, Polar, and Volume objects,
                         respectively.
        ret:
          status       : 0 if save was successful
                         Error code (such as -1,-2) if failed.
        """
        isRightExt = path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        if isRightExt:
            cmd = ("SaveDetector,{:d},{:d},{}"
                   .format(surfaceNumber,objectNumber,fileName))
            reply = self.conversation.Request(cmd)
            return regressLiteralType(reply.rstrip())
        else:
            return -1

    def zSaveFile(self, fileName):
        """Saves the lens currently loaded in the server to a ZEMAX file.

        zSaveFile(fileName)-> status

        args:
          fileName : (string) file name, including full path with extension.
        ret:
          status   :    0 = Zemax successfully saved the lens file & updated the
                            newly saved lens
                     -999 = Zemax couldn't save the file
                       -1 = Incorrect file name
                     Any other value = update failed.

        See also zGetPath, zGetRefresh, zLoadFile, zPushLens.
        """
        isAbsPath = path.isabs(fileName)
        isRightExt = path.splitext(fileName)[1] in ('.zmx','.ZMX')
        if isAbsPath and isRightExt:
            cmd = "SaveFile,{}".format(fileName)
            reply = self.conversation.Request(cmd)
            return int(float(reply.rstrip()))
        else:
            return -1

    def zSaveMerit(self, fileName):
        """Saves the current merit function to a ZEMAX .MF file.

        zSaveMerit(fileName)->meritData

        args:
          fileName  : (string) name of the merit function file with full path and
                       extension.
        ret:
          meritData : If successful, it is the number of operands in the merit
                      function
                      If meritData = -1, saving failed.

        See also: zOptimize, zLoadMerit
        """
        isAbsPath = path.isabs(fileName)
        isRightExt = path.splitext(fileName)[1] in ('.mf','.MF')
        if isAbsPath and isRightExt:
            cmd = "SaveMerit,{}".format(fileName)
            reply = self.conversation.Request(cmd)
            return int(float(reply.rstrip()))
        else:
            return -1

    def zSaveTolerance(self, fileName):
        """Saves the tolerances of the current lens to a file.

        saveTolerance(fileName)->numTolOperands

        args:
          fileName : (string) file name of the file to save the tolerance data.
                     If no path is provided in the filename, the <data>\Tolerance
                     folder is assumed. Although it is not enforced, it is
                     useful to use ".tol" as extension.
        ret:
          numTolOperands : (integer) number of tolerance operands saved.

        See also zLoadTolerance.
        """
        cmd = "SaveTolerance,{}".format(fileName)
        reply = self.conversation.Request(cmd)
        return int(float(reply.rstrip()))

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
        cmd  = ("SetAperture,{sN:d},{aT:d},{aMn:1.20g},{aMx:1.20g},{xD:1.20g},"
                "{yD:1.20g},{aF}".format(sN=surfNum,aT=aType,aMn=aMin,aMx=aMax,
                 xD=xDecenter,yD=yDecenter,aF=apertureFile))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        apertureInfo = tuple([float(elem) for elem in rs])
        return apertureInfo

    def zSetBuffer(self, bufferNumber, textData):
        """Used to store client specific data with the window being created or
        updated. The buffer data can be used to store user selected options,
        instead of using the settings data on the command line of the zMakeTextWindow
        or zMakeGraphicWindow functions. The data must be in a string format.

        zSetBuffer(bufferNumber, textData)->status

        args:
          bufferNumber  : (integer) numbers between 0 and 15 inclusive (for 16
                           buffers provided)
          textData      : (string) is the only text that is stored, maximum of
                          240 characters
        ret:
          status        : 0 if successful, else -1

        Note:
        The buffer data is not associated with any particular window until either
        the  zMakeTextWindow() or zMakeGraphicWindow() functions are issued. Once
        ZEMAX receives the MakeTextWindow or MakeGraphicWindow item, the buffer
        data is then copied to the appropriate window memory, and then may later
        be retrieved from that window's buffer using zGetBuffer() function.

        See also zGetBuffer.
        """
        if (0 < len(textData) < 240) and (0 <= bufferNumber < 16):
            cmd = "SetBuffer,{:d},{}".format(bufferNumber,str(textData))
            reply = self.conversation.Request(cmd)
            return 0 if 'OK' in reply.rsplit() else -1
        else:
            return -1

    def zSetConfig(self,configNumber):
        """Switches the current configuration number (selected column in the MCE),
        and updates the system.

        zSetConfig(configNumber)->(currentConfig, numberOfConfigs, error)

        args:
            configNumber : The configuration (column) number to set current
        ret:
            3-tuple containing the following elements:
             currentConfig    : current configuration (column) number in MCE
                                1 <= currentConfig <= numberOfConfigs
             numberOfConfigs  : number of configs (columns).
             error            : 0  = successful; new current config is traceable
                                -1 = failure

        See also zGetConfig. Use zInsertConfig to insert new configuration in the
        multi-configuration editor.
        """
        reply = self.conversation.Request("SetConfig,{:d}".format(configNumber))
        rs = reply.split(',')
        return tuple([int(elem) for elem in rs])

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
        cmd = ("SetExtra,{:d},{:d},{:1.20g}"
               .format(surfaceNumber,columnNumber,value))
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
        print("Here:",arg1,arg2,arg3)
        if n:
            cmd = ("SetField,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                   ",{:1.20g},{:1.20g},{:1.20g}"
                   .format(n,arg1,arg2,arg3,vdx,vdy,vcx,vcy,van))
        else:
            print("Nowhere:",arg1,arg2,arg3)
            cmd = ("SetField,{:d},{:d},{:d},{:.0f}".format(0,arg1,arg2,arg3))

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
            fieldType       : the field type (0=angle, 1=object height,
                              2=paraxial image height, and 3 = real image height
            fNormalization  : field normalization (0=radial, 1=rectangular)
            iFieldDataTuple : the input field data tuple is an N-D tuple (0<N<=12)
                              with every dimension representing a single field
                              location. It can be constructed as shown here with
                              an example:
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
        cmd = ("SetField,{:d},{:d},{:d},{:d}"
              .format(0,fieldType,fieldCount,fNormalization))
        reply = self.conversation.Request(cmd)
        oFieldDataTuple = [ ]
        for i in range(fieldCount):
            fieldData = self.zSetField(i+1,*iFieldDataTuple[i])
            oFieldDataTuple.append(fieldData)
        return tuple(oFieldDataTuple)

    def zSetFloat(self):
        """Sets all surfaces without surface apertures to have floating apertures.
        Floating apertures will vignette rays which trace beyond the semi-diameter.

        zSetFloat()->status

        args:
            None
        ret:
            status : 0 = success, -1 = fail
        """
        retVal = -1
        reply = self.conversation.Request('SetFloat')
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zSetLabel(self,surfaceNumber,label):
        """This command associates an integer label with the specified surface.
        The label will be retained by ZEMAX as surfaces are inserted or deleted
        around the target surface.

        zSetLabel(surfaceNumber, label)->assignedLabel

        args:
            surfaceNumber : (integer) the surface number
            label         : (integer) the integer label
        rets:
            assignedLabel : (integer) should be equal to label

        See also zGetLabel, zFindLabel
        """
        reply = self.conversation.Request("SetLabel,{:d},{:d}"
                                          .format(surfaceNumber,label))
        return int(float(reply.rstrip()))

    def zSetMulticon(self, config, *multicon_args):
        """Set data or operand type in the multi-configuration editior. Note that
        there are 2 ways of using this function.

        1. If `config` is non-zero, then the function is used to set data in the
           MCE using the following syntax:

        zSetMulticon(config,row,value,status,pickuprow,
                     pickupconfig,scale,offset)->multiConData

        Example: multiConData = zSetMulticon(1,5,5.6,0,0,0,1.0,0.0)

        args:
            config        : (int) configuration number (column)
            row           : (int) row or operand number
            value         : (float) value to set
            status        : (int) see below
            pickuprow     : (int) see below
            pickupconfig  : (int) see below
            scale         : (float)
            offset        : (float)

        ret:
            multiConData is a 8-tuple whose elements are:
            (value,num_config,num_row,status,pickuprow,pickupconfig,scale,offset)

            The status integer is 0 for fixed, 1 for variable, 2 for pickup, and 3
            for thermal pickup. If status is 2 or 3, the pickuprow and pickupconfig
            values indicate the source data for the pickup solve.

        2. If the `config` = 0 , zSetMulticon may be used to set the operand type
           and number data using the following syntax:

        zSetMulticon(0,row,operand_type,number1,number2,number3)-> multiConData

        Example: multiConData = zSetMulticon(0,5,'THIC',15,0,0)

        args:
            config       : 0
            row          : (int) row or operand number in the MCE
            operand_type : (string) operand type, such as 'THIC', 'WLWT', etc.
            number1      : (int)
            number2      : (int)
            number3      : (int)
                           Please refer to "SUMMARY OF MULTI-CONFIGURATION OPERANDS"
                           in the Zemax manual.
        ret:
            multiConData is a 4-tuple whose elements are:
            (operand_type,number1,number2,number3)

        NOTE:
        1. If there are current operands in the MCE, it is recommended to first
           use zInsertMCO to insert a row and then use zSetMulticon(0,...). For
           example use zInsertMCO(5) and then use zSetMulticon(0,5,'THIC',15,0,0).
           If not, then existing rows may be overwritten.
        2. The functions raises an exception if it determines the arguments
           to be invalid.

        See also zGetMulticon()
        """
        if config > 0 and len(multicon_args) == 7:
            (row,value,status,pickuprow,pickupconfig,scale,offset) = multicon_args
            cmd=("SetMulticon,{:d},{:d},{:1.20g},{:d},{:d},{:d},{:1.20g},{:1.20g}"
            .format(config,row,value,status,pickuprow,pickupconfig,scale,offset))
        elif ((config == 0) and (len(multicon_args) == 5) and
                                           (zo.isZOperand(multicon_args[1],3))):
            (row,operand_type,number1,number2,number3) = multicon_args
            cmd=("SetMulticon,{:d},{:d},{},{:d},{:d},{:d}"
            .format(config,row,operand_type,number1,number2,number3))
        else:
            raise ValueError('Invalid input, expecting proper argument')
        # FIX !!! Should it just return -1, instead of raising a value error?
        # If the raise is removed, change code accordingly in the unittest.
        reply = self.conversation.Request(cmd)
        if config: # if config > 0
            rs = reply.split(",")
            multiConData = [float(rs[i]) if (i == 0 or i == 6 or i== 7) else int(rs[i])
                                                 for i in range(len(rs))]
        else: # if config == 0
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0,rs[0])
        return tuple(multiConData)

    def zSetNSCObjectData(self, surfaceNumber, objectNumber, code, data):
        """Sets the various data for NSC objects.

        zSetNSCOjbect(surfaceNumber,objectNumber,code,data)->nscObjectData

        args:
            surfaceNumber : (integer) surface number of the NSC group. Use 1 if
                            the program mode is Non-Sequential.
            objectNumber  : (integer) the NSC ojbect number
            code          : (integer) see the nscObjectData returned table
            data          : data (string/integer/float) to set
        rets:
            nscObjectData : nscObjectData as per the table below if successful
                            else -1
        ------------------------------------------------------------------------
        Code - Data set/returned by SetNSCObjectData (after setting the new data)
        ------------------------------------------------------------------------
          0  - Object type name. (string)
          1  - Comment, which also defines the file name if the object is defined
               by a file. (string)
          2  - Color. (integer)
          5  - Reference object number. (integer)
          6  - Inside of object number. (integer)
        The following codes set values on the Type tab of the Object Properties dialog.
          3  - 1 if object uses a user defined aperture file, 0 otherwise. (integer)
          4  - User defined aperture file name, if any. (string)
         29  - Sets the "Use Pixel Interpolation" checkbox. (1 = checked, 0 = unchecked)
        The following codes set values on the Sources tab of the Object Properties dialog.
        101  - Sets the source object random polarization. (1 = checked, 0 = unchecked)
        102  - Sets the source object reverse rays option. (1 = checked, 0 for unchecked)
        103  - Sets the source object Jones X value.
        104  - Sets the source object Jones Y value.
        105  - Sets the source object Phase X value.
        106  - Sets the source object Phase Y value.
        107  - Sets the source object initial phase in degrees value.
        108  - Sets the source object coherence length value.
        109  - Sets the source object pre-propagation value.
        110  - Sets the source object sampling method; (0 = random, 1 = Sobol sampling)
        111  - Sets the source object bulk scatter method; (0 = many, 1 = once, 2 = never)
        The following codes set values on the Bulk Scatter tab of the Object
        Properties dialog.
        202  - Sets the Mean Path value.
        203  - Sets the Angle value.
        211-226 - Sets the DLL parameter 1-16, respectively.
        ------------------------------------------------------------------------
        See also zSetNSCObjectFaceData
        """
        str_codes = (0,1,4)
        int_codes = (2,3,5,6,29,101,102,110,111)
        if code in str_codes:
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{}"
              .format(surfaceNumber,objectNumber,code,data))
        elif code in int_codes:
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{:d}"
              .format(surfaceNumber,objectNumber,code,data))
        else:  # data is float
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{:1.20g}"
              .format(surfaceNumber,objectNumber,code,data))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscObjectData = -1
        else:
            if code in str_codes:
                nscObjectData = str(rs)
            elif code in int_codes:
                nscObjectData = int(float(rs))
            else:
                nscObjectData = float(rs)
        return nscObjectData

    def zSetNSCObjectFaceData(self,surfNumber,objNumber,faceNumber,code,data):
        """Sets the various data for NSC object faces.

        zSetNSCObjectFaceData(surfNumber,objNumber,faceNumber,code,data)->
                                                                  nscObjFaceData

        args:
          surfNumber   : (integer) surface number. Use 1 if the program mode is
                         Non-Sequential.
          objNumber    : (integer) object number
          faceNumber   : (integer) face number
          code         : (integer) code (see below)
          data         : (float/integer/string) data to set
        ret:
          nscObjFaceData  : data for NSC object faces (see the table for the
                            particular type of data) if successful, else -1
        ------------------------------------------------------------------------
        Code     Data set/returned by zSetNSCObjectFaceData (after setting new data)
        ------------------------------------------------------------------------
         10   -  Coating name. (string)
         20   -  Scatter code. (0 = None, 1 = Lambertian, 2 = Gaussian,
                 3 = ABg, and 4 = user defined.) (integer)
         21   -  Scatter fraction. (double)
         22   -  Number of rays to scatter. (integer)
         23   -  Gaussian scatter sigma. (double)
         24   -  Face is setting;(0 = object default, 1 = reflective,
                 2 = absorbing.) (integer)
         30   -  ABg scatter profile name for reflection. (string)
         31   -  ABg scatter profile name for transmission. (string)
         40   -  User Defined Scatter DLL name. (string)
         41-46 - User Defined Scatter Parameter 1 - 6. (double)
         60   -  User Defined Scatter data file name. (string)
        ------------------------------------------------------------------------
        See also zGetNSCObjectFaceData
        """
        str_codes = (10,30,31,40,60)
        int_codes = (20,22,24)
        if code in str_codes:
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{}"
                   .format(surfNumber,objNumber,faceNumber,code,data))
        elif code in int_codes:
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{:d}"
                  .format(surfNumber,objNumber,faceNumber,code,data))
        else: # data is float
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{:1.20g}"
                  .format(surfNumber,objNumber,faceNumber,code,data))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscObjFaceData = -1
        else:
            if code in str_codes:
                nscObjFaceData = str(rs)
            elif code in int_codes:
                nscObjFaceData = int(float(rs))
            else:
                nscObjFaceData = float(rs)
        return nscObjFaceData

    def zSetNSCParameter(self,surfNumber,objNumber,parameterNumber,data):
        """Sets the parameter data for NSC objects.

        zSetNSCParameter(surfNumber,objNumber,parameterNumber,data)->nscParaVal

        args:
          surfNumber      : (integer) surface number. Use 1 if
                            the program mode is Non-Sequential.
          objNumber       : (integer) object number
          parameterNumber : (integer) parameter number
          data            : (float) new numeric value for the parameterNumber
        ret:
          nscParaVal     : (float) parameter value

        See also zGetNSCParameter
        """
        cmd = ("SetNSCParameter,{:d},{:d},{:d},{:1.20g}"
              .format(surfNumber,objNumber,parameterNumber,data))
        reply = self.conversation.Request(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zSetNSCPosition(self,surfNumber,objectNumber,code,data):
        """Returns the position data for NSC objects.

        zSetNSCPosition(surfNumber,objectNumber,code,data)->nscPosData

        arg:
          surfNumber   : (integer) surface number. Use 1 if
                         the program mode is Non-Sequential.
          objectNumber : (integer) object number
          code         : (integer) 1-7 for x, y, z, tilt-x, tilt-y, tilt-z, and
                         material, respectively.
          data         : numeric (float) for codes 1-6, string for material (code-7)
        ret:
          nscPosData is a 7-tuple containing x,y,z,tilt-x,tilt-y,tilt-z,material

        See also zGetNSCPosition
        """
        if code == 7:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{}"
            .format(surfNumber,objectNumber,code,data))
        else:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{:1.20g}"
            .format(surfNumber,objectNumber,code,data))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPosData = -1
        else:
            nscPosData = tuple([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                    for i in range(len(rs))])
        return nscPosData

    def zSetNSCProperty(self,surfaceNumber,objectNumber,faceNumber,code,value):
        """Sets a numeric or string value to the property pages of objects
        defined in the non-sequential components editor. It mimics the ZPL
        function NPRO.

        zSetNSCProperty(surfaceNumber,objectNumber,faceNumber,code,value)->nscPropData

        args:
          surfaceNumber  : (integer) surface number. Use 1 if the program mode is
                           Non-Sequential.
          objectNumber   : (integer) object number
          faceNumber     : (integer) face number
          code           : (integer) code used to identify specific types of NSC
                           object data (see below)
          value          : (string, float or integer) value to set (see below)
        ret:
          nscPropData  :  string or numeric (see below)
        ------------------------------------------------------------------------
        Code   Property
        ------------------------------------------------------------------------
        The following codes set values on the NSC Editor.
          1 - Sets the object comment. (string)
          2 - Sets the reference object number. (integer)
          3 - Sets the "inside of" object number. (integer)
          4 - Sets the object material. (string)
        The following codes set values on the Type tab of the Object Properties
        dialog.
          0 - Sets the object type. The value should be the name of the object,
              such as "NSC_SLEN" for the standard lens. The names for each object
              type are listed in the Prescription Report for each object type in
              the NSC editor. All NSC object names start with "NSC_". (string)
         13 - Sets User Defined Aperture, (1 = checked, 0 = unchecked)
         14 - Sets the User Defined Aperture file name. (string)
         15 - Sets the "Use Global XYZ Rotation Order" checkbox, (1 = checked,
              0 = unchecked). (integer)
         16 - Sets the "Rays Ignore This Object" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
         17 - Sets the "Object Is A Detector" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
         18 - Sets the "Consider Objects" list. The argument should be a string
              listing the object numbers to consider delimited by spaces, such
              as "2 5 14". (string)
         19 - Sets the "Ignore Objects" list. The argument should be a string
              listing the object numbers to ignore delimited by spaces, such as
              "1 3 7". (string)
         20 - Sets the "Use Pixel Interpolation" checkbox, (1 = checked, 0 = un-
              checked.) (integer)
        The following codes set values on the Coat/Scatter tab of the Object Pro-
        perties dialog.
          5 - Sets the coating name for the specified face. (string)
          6 - Sets the profile name for the specified face. (string)
          7 - Sets the scatter mode for the specified face, (0 = none,
              1 = Lambertian, 2 = Gaussian, 3 = ABg, 4 = User Defined.)
          8 - Sets the scatter fraction for the specified face. (float)
          9 - Sets the number of scatter rays for the specified face. (integer)
         10 - Sets the Gaussian sigma for the specified face. (float)
         11 - Sets the reflect ABg data name for the specified face. (string)
         12 - Sets the transmit ABg data name for the specified face. (string)
         27 - Sets the name of the user defined scattering DLL. (string)
         28 - Sets the name of the user defined scattering data file. (string)
        21-26 - Sets parameter values on the user defined scattering DLL. (float)
         29 - Sets the "Face Is" property for the specified face. (0 = "Object
              Default", 1 = "Reflective", 2 = "Absorbing")
        The following codes set values on the Bulk Scattering tab of the Object
        Properties dialog.
         81 - Sets the "Model" value on the bulk scattering tab. (0 = "No Bulk
              Scattering", 1 = "Angle Scattering", 2 = "DLL Defined Scattering")
         82 - Sets the mean free path to use for bulk scattering.
         83 - Sets the angle to use for bulk scattering.
         84 - Sets the name of the DLL to use for bulk scattering.
         85 - Sets the parameter value to pass to the DLL, where the face value
              is used to specify which parameter is being defined. The first
              parameter is 1, the second is 2, etc. (float)
         86 - Sets the wavelength shift string. (string)
        The following codes set values on the Diffraction tab of the Object Pro-
        perties dialog.
         91 - Sets the "Split" value on the diffraction tab. (0 = "Don't Split
              By Order", 1 = "Split By Table Below", 2 = "Split By DLL Function")
         92 - Sets the name of the DLL to use for diffraction splitting. (string)
         93 - Sets the Start Order value. (float)
         94 - Sets the Stop Order value. (float)
         95 - Sets  the  parameter  values  on  the  diffraction  tab.  These
              are  the  parameters passed to the diffraction splitting DLL as
              well as the order efficiency values used by "split by table below"
              option. The face value is used to specify which parameter is being
              defined. The first parameter is 1, the second is 2, etc. (float)
        The following codes set values on the Sources tab of the Object Properties
        dialog.
        101 - Sets the source object random polarization.(1=checked,0=unchecked)
        102 - Sets the source object reverse rays option.(1=checked,0=unchecked)
        103 - Sets the source object Jones X value.
        104 - Sets the source object Jones Y value.
        105 - Sets the source object Phase X value.
        106 - Sets the source object Phase Y value.
        107 - Sets the source object initial phase in degrees value.
        108 - Sets the source object coherence length value.
        109 - Sets the source object pre-propagation value.
        110 - Sets the source object sampling method; (0=random,1=Sobol sampling)
        111 - Sets the source object bulk scatter method; (0=many,1=once,2=never)
        112 - Sets the array mode; (0 = none, 1 = rectangular, 2 = circular,
              3 = hexapolar, 4 = hexagonal)
        113 - Sets the source color mode. For a complete list of the available
              modes, see "Defining the color and spectral content of sources"
              in the Zemax manual. The source color modes are numbered starting
              with 0 for the System Wavelengths, and then from 1 through the last
              model listed in the dialog box control. (integer)
        114-116 - Sets the number of spectrum steps, start wavelength, and end
                  wavelength, respectively. (float)
        117 - Sets the name of the spectrum file. (string)
        161-162 - Sets the array mode integer arguments 1 and 2.
        165-166 - Sets the array mode double precision arguments 1 and 2.
        181-183 - Sets the source color mode arguments, for example, the XYZ
                  values of the Tristimulus. (float)
        The following codes set values on the Grin tab of the Object Properties
        dialog.
        121 - Sets the "Use DLL Defined Grin Media" checkbox. (1 = checked, 0 =
              unchecked) (integer)
        122 - Sets the Maximum Step Size value. (float)
        123 - Sets the DLL name. (string)
        124 - Sets the Grin DLL parameters. These are the parameters passed to
              the DLL. The face value is used to specify which parameter is being
              defined. The first parameter is 1, the second is 2, etc. (float)
        The following codes set values on the Draw tab of the Object Properties
        dialog.
        141 - Sets the do not draw object checkbox.(1 = checked, 0 = unchecked)
        142 - Sets the object opacity. (0 = 100%, 1 = 90%, 2 = 80%, etc.)
        The following codes set values on the Scatter To tab of the Object
        Properties dialog.
        151 - Sets the scatter to method. (0 = scatter to list, 1 = importance
              sampling)
        152 - Sets  the  Importance  Sampling  target  data.  The  argument
              should be a string listing the ray number, the object number, the
              size, & the limit value, all separated by spaces Here is a sample
              syntax to set the Importance Sampling data for ray 3, object 6,
              size 3.5, and limit 0.6:"3 6 3.5 0.6". (string)
        153 - Sets the "Scatter To List" values. The argument should be a string
              listing the object numbers to scatter to delimited by spaces, such
              as "4 6 19". (string)
        The following codes set values on the Birefringence tab of the Object
        Properties dialog.
        171 - Sets the Birefringent Media checkbox. (0 = unchecked, 1 = checked)
        172 - Sets the Birefringent Media Mode. (0 = Trace ordinary and
              extraordinary rays, 1 = Trace only ordinary rays, 2 = Trace only
              extraordinary rays, and 3 = Waveplate mode) (integer)
        173 - Sets the Birefringent Media Reflections status. (0 = Trace
              reflected and refracted rays, 1 = Trace only refracted rays, and
              2 = Trace only reflected rays)
        174-176 - Sets the Ax, Ay, and Az values. (float)
        177 - Sets the Axis Length. (float)
        ------------------------------------------------------------------------
        See also zGetNSCProperty
        """
        cmd = ("SetNSCProperty,{:d},{:d},{:d},{:d},"
                .format(surfaceNumber,objectNumber,code,faceNumber))
        if code in (0,1,4,5,6,11,12,14,18,19,27,28,84,86,92,117,123):
            cmd = cmd + value
        elif code in (2,3,7,9,13,15,16,17,20,29,81,91,101,102,110,111,113,121,
                                       141,142,151,152,153161,162,171,172,173):
            cmd = cmd + str(int(value))
        else:
            cmd = cmd + str(float(value))
        reply = self.conversation.Request(cmd)
        nscPropData = process_get_set_NSCProperty(code,reply)
        return nscPropData

    def zSetNSCSettings(self,nscSettingsData):
        """Sets the maximum number of intersections, segments, nesting level,
        minimum absolute intensity, minimum relative intensity, glue distance,
        miss ray distance, and ignore errors flag used for NSC ray tracing.

        zSetNSCSettings(nscSettingsData)->nscSettingsDataRet

        args:
          nscSettingsData is an 8-tuple with the following elements
          maxInt     : (integer) maximum number of intersections
          maxSeg     : (integer) maximum number of segments
          maxNest    : (integer) maximum nesting level
          minAbsI    : (float) minimum absolute intensity
          minRelI    : (float) minimum relative intensity
          glueDist   : (float) glue distance
          missRayLen : (float) miss ray distance
          ignoreErr  : (integer) 1 if true, 0 if false
        rets:
          nscSettingsDataRet is also an 8-tuple with the same elements as
          nscSettingsData.

        NOTE:
             Since the `maxSeg` value may require large amounts of RAM, verify
             that the new value was accepted by checking the returned tuple.

        See also zGetNSCSettings
        """
        (maxInt,maxSeg,maxNest,minAbsI,minRelI,glueDist,missRayLen,
                                                 ignoreErr) = nscSettingsData
        cmd = ("SetNSCSettings,{:d},{:d},{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:d}"
        .format(maxInt,maxSeg,maxNest,minAbsI,minRelI,glueDist,missRayLen,ignoreErr))
        reply = str(self.conversation.Request(cmd))
        rs = reply.rsplit(",")
        nscSettingsData = [float(rs[i]) if i in (3,4,5,6) else int(float(rs[i]))
                                                        for i in range(len(rs))]
        return tuple(nscSettingsData)

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
        cmd = "SetWave,{:d},{:d},{:d}".format(0,primaryWaveNumber,waveData[1])
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        waveData = tuple([int(elem) for elem in rs])
        return waveData

    def zSetOperand(self, row, column, value):
        """Sets the operand data in the Merit Function Editor.

        zSetOperand(row,column,value)->operandData

        args:
          row    : (integer) row operand number in the MFE
          column : (integer) column number (see table below)
                        -----------------------------------------------
                        Column #           value
                        -----------------------------------------------
                         1               operand type (string)
                         2               int1 (integer)
                         3               int2 (integer)
                         4-7             data1-data4 (float)
                         8               target (float)
                         9               weight (float)
                         10              contribution (float)
                         12-13           data5-data6 (float)
                        -----------------------------------------------
          value : string/integer/float. See table above.
        ret:
          operandData : the value (string/integer/float) set in the MFE cell

        Note:
            1. To update the merit function after called zSetOperand() function,
               use the zOptimize() function with the number of cycles set to -1.
            2. Use zInsertMFO() to insert additional rows, before calling
               zSetOperand().

        See also zGetOperand, zOptimize, zInsertMFO.
        """
        if column == 1:
            if zo.isZOperand(str(value)):
                value = str(value)
            else:
                print("Not a valid operand in zSetOperand().")
                return -1
        elif column in (2,3):
            value = '{}'.format(int(float(value)))
        else:
            value = '{}'.format(float(value))
        cmd = "SetOperand,{:d},{:d},{}".format(row, column, value)
        reply = self.conversation.Request(cmd)
        return process_get_set_Operand(column, reply)

    def zSetPolState(self,nlsPolarized,Ex,Ey,Phx,Phy):
        """Sets the default polarization state. These parameters correspond to
        the Polarization tab under the General settings.

        zSetPolState()->polStateData

        args:
          nlsPolarized : (integer) if nlsPolarized > 0, then default polarization
                         state is unpolarized.
          Ex           : (float) normalized electric field magnitude in x direction
          Ey           : (float) normalized electric field magnitude in y direction
          Phax         : (float) relative phase in x direction in degrees
          Phay         : (float) relative phase in y direction in degrees

        ret:
          polStateData is 5-tuple containing nlsPolarized, Ex, Ey, Phax, and Phay

        Note: The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
              values are accepted.

        See also zGetPolState.
        """
        cmd = ("SetPolState,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                .format(nlsPolarized,Ex,Ey,Phx,Phy))
        reply = self.conversation.Request(cmd)
        rs = reply.rsplit(",")
        polStateData = [int(float(elem)) if i==0 else float(elem)
                                       for i,elem in enumerate(rs[:-1])]
        return tuple(polStateData)

    def zSetSettingsData(self, number, data):
        """This function sets the settings data used by a window in temporary
        storage before calling zMakeGraphicWindow or zMakeTextWindow. The data
        may be retrieved using zGetSettingsData.

        zSetSettingsData(number, data)->settingsData

        args:
          number  : (integer) Currently, only number = 0 is supported. This number
                    may be used to expand the feature in the future.
          data    :
        ret:
          settingsData: (string)

        Note: Please refer to "How ZEMAX calls the client" in the Zemax manual.
        See also zGetSettingsData.
        """
        cmd = "SettingsData,{:d},{}".format(number,data)
        reply = self.conversation.Request(cmd)
        return str(reply.rstrip())

    def zSetSolve(self, surfaceNumber, code, *solveData):
        """Sets data for solves and/or pickups on the surface with number
        `surfaceNumber`.

        zSetSolve(surfaceNumber, code, solveData)->solveData

        args:
          surfaceNumber : (integer) surface number for which the solve is to be
                          set.
          code          : (integer) indicating which surface parameter the solve
                           data is for, such as curvature, thickness, glass, conic,
                           semi-diameter, etc. (see the table below)
          *solveData    : is a set of arguments (could be a tuple) containing
                          `solvetype`, `param1`,`param2`,`param3`,and `param4`,
                          depending on the code value according to the
                          following table.
        ------------------------------------------------------------------------
        GetSolve Code           -  Solve data (solveData) & Returned data format
        ------------------------------------------------------------------------
        0 (curvature)           -  solvetype, parameter1, parameter2, pickupcolumn
        1 (thickness)           -  solvetype, parameter1, parameter2, parameter3,
                                   pickupcolumn
        2 (glass)               -  solvetype (for solvetype = 0)
                                   solvetype, Index, Abbe, Dpgf (for solvetype = 1,
                                   model glass)
                                   solvetype, pickupsurf (for solvetype = 2, pickup)
                                   solvetype, index_offset, abbe_offset (for
                                   solvetype = 4, offset)
                                   solvetype (for solvetype = all other values)
        3 (semi-diameter)        - solvetype, pickupsurf, pickupcolumn
        4 (conic)                - solvetype, pickupsurf, pickupcolumn
        5-16 (parameters 1-12)   - solvetype, pickupsurf, offset,  scalefactor,
                                   pickupcolumn
        17 (parameter 0)         - solvetype, pickupsurf, offset,  scalefactor,
                                   pickupcolumn
        1001+ (extra data values 1+) - solvetype, pickupsurf, scalefactor, offset,
                                       pickupcolumn
        ret:
          solveData     : a tuple, depending on the code value according to the
                          above table (same return as zGetSolve), if successful,
                          -1 if the command is a 'BAD COMMAND'

        Note: The `solvetype` is an integer code, & the parameters have meanings
        that depend upon the solve type; see the chapter "SOLVES" in the Zemax
        manual for details.

        See also zGetSolve, zGetNSCSolve, zSetNSCSolve.
        """
        if not solveData:
            print("Error [zSetSolve] No solve data passed.")
            return -1
        try:
            if code == 0:                    # Solve specified on CURVATURE
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # marninal ray angle (M)
                    data = '{:1.20g}'.format(solveData[1]) # angle
                elif solveData[0] == 3:         # chief ray angle (C)
                    data = '{:1.20g}'.format(solveData[1]) # angle
                elif solveData[0] == 4:         # pickup (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3])) # suface,scale-factor,column
                elif solveData[0] == 5:         # marginal ray normal (N)
                    data = ''
                elif solveData[0] == 6:         # chief ray normal (N)
                    data = ''
                elif solveData[0] == 7:         # aplanatic (A)
                    data = ''
                elif solveData[0] == 8:         # element power (X)
                    data = '{:1.20g}'.format(solveData[1]) # power
                elif solveData[0] == 9:         # concentric with surface (S)
                    data = '{:d}'.format(solveData[1]) # surface to be concentric to
                elif solveData[0] == 10:        # concentric with radius (R)
                    data = '{:d}'.format(solveData[1]) # surface to be concentric with
                elif solveData[0] == 11:        # f/# (F)
                    data = '{:1.20g}'.format(solveData[1]) # paraxial f/#
                elif solveData[0] == 12:        # zpl macro (Z)
                    data = str(solveData[1])       # macro name
            elif code == 1:                  # Solve specified on THICKNESS
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # marninal ray height (M)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1],solveData[2]) # height, pupil zone
                elif solveData[0] == 3:         # chief ray height (C)
                    data = '{:1.20g}'.format(solveData[1])   # height
                elif solveData[0] == 4:         # edge thickness (E)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1],solveData[2]) # thickness, radial height (0 for semi-diameter)
                elif solveData[0] == 5:         # pickup (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3],solveData[4])) # surface, scale-factor, offset, column
                elif solveData[0] == 6:         # optical path difference (O)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1],solveData[2]) # opd, pupil zone
                elif solveData[0] == 7:         # position (T)
                    data = '{:d},{:1.20g}'.format(solveData[1],solveData[2]) # surface, length from surface
                elif solveData[0] == 8:         # compensator (S)
                    data = '{:d},{:1.20g}'.format(solveData[1],solveData[2]) # surface, sum of surface thickness
                elif solveData[0] == 9:         # center of curvature (X)
                    data = '{:d}'.format(solveData[1]) # surface to be at the COC of
                elif solveData[0] == 10:        # pupil position (U)
                    data = ''
                elif solveData[0] == 11:        # zpl macro (Z)
                    data = str(solveData[1])       # macro name
            elif code == 2:                  # Solve specified on GLASS
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # model
                    data = ('{:1.20g},{:1.20g},{:1.20g}'
                    .format(solveData[1],solveData[2],solveData[3])) # index Nd, Abbe Vd, Dpgf
                elif solveData[0] == 2:         # pickup (P)
                    data = '{:d}'.format(solveData[1]) # surface
                elif solveData[0] == 3:         # substitute (S)
                    data = str(solveData[1])      # catalog name
                elif solveData[0] == 4:         # offset (O)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1],solveData[2]) # index Nd offset, Abbe Vd offset
            elif code == 3:                  # Solve specified on SEMI-DIAMETER
                if solveData[0] == 0:           # automatic
                    data = ''
                elif solveData[0] == 1:         # fixed (U)
                    data = ''
                elif solveData[0] == 2:         # pickup (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3])) # surface, scale-factor, column
                elif solveData[0] == 3:         # maximum (M)
                    data = ''
                elif solveData[0] == 4:         # zpl macro (Z)
                    data = str(solveData[1])       # macro name
            elif code == 4:                  # Solve specified on CONIC
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # pickup (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3])) # surface, scale-factor, column
                elif solveData[0] == 3:         # zpl macro (Z)
                    data = str(solveData[1])       # macro name
            elif code in range(5,17):        # Solve specified on PARAMETERS 1-12
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # pickup (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3],solveData[4])) # surface, scale-factor, offset, column
                elif solveData[0] == 3:         # chief ray (C)
                    data = '{:d},{:1.20g}'.format(solveData[1],solveData[2]) # field, wavelength
                elif solveData[0] == 4:         # zpl macro (Z)
                    data = str(solveData[1])       # macro name
            elif code == 17:                 # Solve specified on PARAMETER 0
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # pickup (P)
                    data = '{:d}'.format(solveData[1]) # surface
            elif code > 1000:                # Solve specified on EXTRA DATA VALUES
                if solveData[0] == 0:           # fixed
                    data = ''
                elif solveData[0] == 1:         # variable (V)
                    data = ''
                elif solveData[0] == 2:         # pickup (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3],solveData[4])) # surface, scale-factor, offset, column
                elif solveData[0] == 3:         # zpl macro (Z)
                    data = str(solveData[1])       # macro name
        except IndexError:
            print("Error [zSetSolve]: Check number of solve parameters!")
            return -1
        #synthesize the command to pass to zemax
        if data:
            cmd = ("SetSolve,{:d},{:d},{:d},{}"
                  .format(surfaceNumber,code,solveData[0],data))
        else:
            cmd = ("SetSolve,{:d},{:d},{:d}"
                  .format(surfaceNumber,code,solveData[0]))
        reply = self.conversation.Request(cmd)
        solveData = process_get_set_Solve(reply)
        return solveData

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
        ------------------------------------------------------------------------
        Code      - Datum to be set by by zSetSurfaceData
        ------------------------------------------------------------------------
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
        ------------------------------------------------------------------------
        See also zGetSurfaceData and ZemaxSurfTypes
        """
        cmd = "SetSurfaceData,{:d},{:d}".format(surfaceNumber,code)
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
        cmd = ("SetSurfaceParameter,{:d},{:d},{:1.20g}"
               .format(surfaceNumber,parameter,value))
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
        cmd = ("SetSystem,{:d},{:d},{:d},{:d},{:1.20g},{:1.20g},{:d}"
              .format(unitCode,stopSurf,rayAimingType,useEnvData,temp,pressure,
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
        cmd = ("SetSystemAper,{:d},{:d},{:1.20g}"
               .format(aType,stopSurf,apertureValue))
        reply = self.conversation.Request(cmd)
        rs = reply.split(',')
        systemAperData = tuple([float(elem) if i==2 else int(float(elem))
                                for i, elem in enumerate(rs)])
        return systemAperData

    def zSetSystemProperty(self, code, value1, value2=0):
        """Sets system properties of the system.

        System properties such as system aperture, field, wavelength, and other data,
        based on the integer `code` passed.

        zSetSystemProperty(code)-> sysPropData
        args:
            code        : (integer) value that defines the specific system property
                          to be set (see below).
            value1      : (integer/float/string) depending on `code`
            value2      : (integer/float), ignored if not used
        ret:
            sysPropData : Returned system property data. Either a string or numeric
                          data.

        This function mimics the ZPL function SYPR.
        ------------------------------------------------------------------------
        Code    Property (the values in the bracket are the expected returns)
        ------------------------------------------------------------------------
          4   - Adjust Index Data To Environment. (0:off, 1:on.)
         10   - Aperture Type code. (0:EPD, 1:IF/#, 2:ONA, 3:FBS, 4:PWF/#, 5:OCA)
         11   - Aperture Value. (stop surface semi-diameter if aperture type is FBS,
                else system aperture)
         12   - Apodization Type code. (0:uniform, 1:Gaussian, 2:cosine cubed)
         13   - Apodization Factor.
         14   - Telecentric Object Space. (0:off, 1:on)
         15   - Iterate Solves When Updating. (0:off, 1:on)
         16   - Lens Title.
         17   - Lens Notes.
         18   - Afocal Image Space. (0:off or "focal mode", 1:on or "afocal mode")
         21   - Global coordinate reference surface.
         23   - Glass catalog list. (Use a string or string variable with the glass
                catalog name, such as "SCHOTT". To specify multiple catalogs use
                a single string or string variable containing names separated by
                spaces, such as "SCHOTT HOYA OHARA".)
         24   - System Temperature in degrees Celsius.
         25   - System Pressure in atmospheres.
         26   - Reference OPD method. (0:absolute, 1:infinity, 2:exit pupil, 3:absolute 2.)
         30   - Lens Units code. (0:mm, 1:cm, 2:inches, 3:Meters)
         31   - Source Units Prefix. (0:Femto, 1:Pico, 2:Nano, 3:Micro, 4:Milli,
                5:None,6:Kilo, 7:Mega, 8:Giga, 9:Tera)
         32   - Source Units. (0:Watts, 1:Lumens, 2:Joules)
         33   - Analysis Units Prefix. (0:Femto, 1:Pico, 2:Nano, 3:Micro, 4:Milli,
                5:None,6:Kilo, 7:Mega, 8:Giga, 9:Tera)
         34   - Analysis Units "per" Area. (0:mm^2, 1:cm^2, 2:inches^2, 3:Meters^2, 4:feet^2)
         35   - MTF Units code. (0:cycles per millimeter, 1:cycles per milliradian.
         40   - Coating File name.
         41   - Scatter Profile name.
         42   - ABg Data File name.
         43   - GRADIUM Profile name.
         50   - NSC Maximum Intersections Per Ray.
         51   - NSC Maximum Segments Per Ray.
         52   - NSC Maximum Nested/Touching Objects.
         53   - NSC Minimum Relative Ray Intensity.
         54   - NSC Minimum Absolute Ray Intensity.
         55   - NSC Glue Distance In Lens Units.
         56   - NSC Missed Ray Draw Distance In Lens Units.
         57   - NSC Retrace Source Rays Upon File Open. (0:no, 1:yes)
         58   - NSC Maximum Source File Rays In Memory.
         59   - Simple Ray Splitting. (0:no, 1:yes)
         60   - Polarization Jx.
         61   - Polarization Jy.
         62   - Polarization X-Phase.
         63   - Polarization Y-Phase.
         64   - Convert thin film phase to ray equivalent. (0:no, 1:yes)
         65   - Unpolarized. (0:no, 1:yes)
         66   - Method. (0:X-axis, 1:Y-axis, 2:Z-axis)
         70   - Ray Aiming. (0:off, 1:on, 2:aberrated)
         71   - Ray aiming pupil shift x.
         72   - Ray aiming pupil shift y.
         73   - Ray aiming pupil shift z.
         74   - Use Ray Aiming Cache. (0:no, 1:yes)
         75   - Robust Ray Aiming. (0:no, 1:yes)
         76   - Scale Pupil Shift Factors By Field. (0:no, 1:yes)
         77   - Ray aiming pupil compress x.
         78   - Ray aiming pupil compress y.
         100  - Field type code. (0=angl,1=obj ht,2=parx img ht,3=rel img ht)
         101  - Number of fields.
         102  - The field number is value1, value2 is the field x coordinate
         103  - The field number is value1, value2 is the field y coordinate
         104  - The field number is value1, value2 is the field weight
         105  - The field number is value1, value2 is the field vignetting decenter x
         106  - The field number is value1, value2 is the field vignetting decenter y
         107  - The field number is value1, value2 is the field vignetting compression x
         108  - The field number is value1, value2 is the field vignetting compression y
         109  - The field number is value1, value2 is the field vignetting angle
         110  - The field normalization method, value 1 is 0 for radial and 1 for rectangular
         200  - Primary wavelength number.
         200  - Primary wavelength number.
         201  - Number of wavelengths
         202  - The wavelength number is value1, value 2 is the wavelength in micrometers.
         203  - The wavelength number is value1, value 2 is the wavelength weight
         901  - The number of CPU's to use in multi-threaded computations, such as
                optimization. (0=default). See the manual for details.
        ------------------------------------------------------------------------
        NOTE: Currently Zemax returns just "0" for the codes: 102,103, 104,105,
              106,107,108,109, and 110. This is unexpected! So, PyZDDE will return
              the reply (string) as is for the user to handle. The zSetSystemProperty
              functions as expected nevertheless.

        See also zGetSystemProperty.
        """
        cmd = "SetSystemProperty,{c:d},{v1},{v2}".format(c=code,v1=value1,v2=value2)
        reply = self.conversation.Request(cmd)
        sysPropData = process_get_set_SystemProperty(code,reply)
        return sysPropData

    def zSetTol(self,operandNumber,col,value):
        """Sets the tolerance operand data.

        zSetTol(operandNumber,col,value)-> toleranceData

        args:
          operandNumber : (integer) tolerance operand number (row number in the
                          tolerance editor, when greater than 0)
          col           : (integer) 1 for tolerance Type
                          (integer) 2-4 for int1-int3
                          (integer) 5 for min
                          (integer) 6 for max
          value         : 4-character string (tolerancing operand code) if col==1,
                          else float value to set
        ret:
          toleranceData. It is a number or a 6-tuple, depending upon `operandNumber`
          as follows:
            if operandNumber == 0, toleranceData = number where `number` is the number of
            tolerance operands defined.
            if operandNumber > 0, toleranceData = (tolType, int1, int2, min, max, int3)

          it returns -1 if an error occurs.

        See also zSetTolRow, zGetTol,
        """
        if col == 1: # value is string code for the operand
            if zo.isZOperand(str(value),2):
                cmd = "SetTol,{:d},{:d},{}".format(operandNumber,col,value)
            else:
                return -1
        else:
            cmd = "SetTol,{:d},{:d},{:1.20g}".format(operandNumber,col,value)
        reply = self.conversation.Request(cmd)
        if operandNumber == 0: # returns just the number of operands
            return int(float(reply.rstrip()))
        else:
            return process_get_set_Tol(operandNumber,reply)
        # FIX !!! currently, I am not able to set more than 1 row in the tolerance
        # editor, through this command. I don't find anything like zInsertTol ...
        # A similar function exist for Multi-Configuration editor (zInsertMCO) and
        # for Multi-function editor (zInsertMFO). May need to contact Zemax Support.

    def zSetTolRow(self,operandNumber,tolType,int1,int2,int3,minT,maxT):
        """Helper function to set all the elements of a row (given by operandNumber)
        in the tolerance editor.

        zSetTolRow(operandNumber,tolType,int1,int2,int3,minT,maxT)->tolData

        args:
          operandNumber  : (integer greater than 0) tolerance operand number (row
                           number in the tolerance editor)
          tolType        : 4-character string (tolerancing operand code)
          int1           : (integer) int1 parameter
          int2           : (integer) int2 parameter
          int3           : (integer) int3 parameter
          minT           : (float) minimum value
          maxT           : (float) maximum value
        ret:
          tolData        : tolerance data for the row indicated by the operandNumber
                           if successful, else -1
        """
        tolData = self.zSetTol(operandNumber,1,tolType)
        if tolData != -1:
            self.zSetTol(operandNumber,2,int1)
            self.zSetTol(operandNumber,3,int2)
            self.zSetTol(operandNumber,4,int3)
            self.zSetTol(operandNumber,5,minT)
            self.zSetTol(operandNumber,6,maxT)
            return self.zGetTol(operandNumber)
        else:
            return -1

    def zSetUDOItem(self, bufferCode, dataNumber, data):
        """This function is used to pass just one datum computed by the client
        program to the ZEMAX optimizer. The only time this item name should be
        used is when implementing a User Defined Operand, or UDO. UDO's are
        described in "Optimizing with externally compiled programs" in the Zemax
        manual.

        zGetUDOItem(bufferCode dataNumber, data)->

        args:
          bufferCode : (integer) The buffercode is an integer value provided
                       by ZEMAX to the client that uniquely identifies the
                       correct lens.
          dataNumber : (integer)
          data       : (float) data item number being passed
        ret:
          ?

        Note: After the last data item has been sent, the buffer must be closed
        using the zCloseUDOData() function before the optimization may proceed.
        A typical implementation may consist of the following series of function
        calls:
            zSetUDOItem(bufferCode, 0, value0)
            zSetUDOItem(bufferCode, 1, value1)
            zSetUDOItem(bufferCode, 2, value2)
            zCloseUDOData(bufferCode)

        See also zGetUDOSystem, zCloseUDOData.
        """
        cmd = "SetUDOItem,{:d},{:d},{:1.20g}".format(bufferCode,dataNumber,data)
        reply = self.conversation.Request(cmd)
        return regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zSetVig(self):
        """Request Zemax to set the vignetting factors automatically. Calling this
        function is equivalent to clicking the "Set Vig" button from the "Field Data"
        window. For more information on how Zemax calculates the vignetting factors
        automatically, please refer to "Vignetting factors" under the "Systems Menu"
        chapter in the Zemax Manual.

        zSetVig()->retVal

        args:
          None
        ret:
          retVal  : 0 = success, -1 = fail
        """
        retVal = -1
        reply = self.conversation.Request("SetVig")
        if 'OK' in reply.split():
            retVal = 0
        return retVal

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
            cmd = "SetWave,{:d},{:1.20g},{:1.20g}".format(n,arg1,arg2)
        else:
            cmd = "SetWave,{:d},{:d},{:d}".format(0,arg1,arg2)

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
            cmd = ("SetWave,{:d},{:1.20g},{:1.20g}"
                   .format(i+1,iWaveDataTuple[0][i],iWaveDataTuple[1][i]))
            reply = self.conversation.Request(cmd)
            rs = reply.split(',')
            oWaveDataTuple[0].append(float(rs[0])) # store the wavelength
            oWaveDataTuple[1].append(float(rs[1])) # store the weight
        return (tuple(oWaveDataTuple[0]),tuple(oWaveDataTuple[1]))

    def zWindowMaximize(self,windowNumber=0):
        """Maximize the main ZEMAX window or any analysis window ZEMAX currently
        displayed.

        zWindowMaximize(windowNumber)->retVal

        args:
          windowNumber  : (integer) window number. use 0 for the main ZEMAX window
        ret:
          retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self.conversation.Request("WindowMaximize,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowMinimize(self,windowNumber=0):
        """Minimize the main ZEMAX window or any analysis window ZEMAX currently
        displayed.

        zWindowMinimize(windowNumber)->retVal

        args:
          windowNumber  : (integer) window number. use 0 for the main ZEMAX window
        ret:
          retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self.conversation.Request("WindowMinimize,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowRestore(self,windowNumber=0):
        """Restore the main ZEMAX window or any analysis window to it's previous
        size and position.

        zWindowRestore(windowNumber)->retVal

        args:
          windowNumber  : (integer) window number. use 0 for the main ZEMAX window
        ret:
          retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self.conversation.Request("WindowRestore,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

# ****************************************************************
#                      CONVENIENCE FUNCTIONS
# ****************************************************************
    def spiralSpot(self,hy,hx,waveNum,spirals,rays,mode=0):
        """Convenience function to produce a series of x,y values of rays traced
        in a spiral over the entrance pupil to the image surface. i.e. the final
        destination of the rays is the image surface. This function imitates its
        namesake from MZDDE toolbox (Note: unlike the spiralSpot of MZDDE, you
        are not required to call zLoadLens() before calling spiralSpot()).

        spiralSpot(hy,hx,waveNum,spirals,rays[,mode])->(x,y,z,intensity)

        args:
          hx      : normalized field height along x axis
          hy      : normalized field height along y axis
          waveNum : wavelength number as in the wavelength data editor
          mode    : 0 = real, 1 = paraxial
        ret:
          rayInfo : 4-tuple = (x,y,z,intensity)
        """
        # Calculate the ray pattern on the pupil plane
        finishAngle = spirals*2*pi
        dTheta = finishAngle/(rays-1)
        theta = [i*dTheta for i in range(rays)]
        r = [i/finishAngle for i in theta]
        pXY = [(ri*cos(thetai), ri*sin(thetai)) for ri, thetai in izip(r,theta)]
        x = [] # x-coordinate of the image surface
        y = [] # y-coordinate of the image surface
        z = [] # z-coordinate of the image surface
        intensity = [] # the relative transmitted intensity of the ray
        for px,py in pXY:
            rayTraceData = self.zGetTrace(waveNum,mode,-1,hx,hy,px,py)
            if rayTraceData[0] == 0:
                x.append(rayTraceData[2])
                y.append(rayTraceData[3])
                z.append(rayTraceData[4])
                intensity.append(rayTraceData[11])
            else:
                print("Raytrace Error")
                exit()
                # !!! FIX raise an error here
        return (x,y,z,intensity)


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
            ##print(Warning: Scaling of aperture may be incorrect)
            pass
        elif sysAperData[0] == 3: # System aperture if float by stop
            pass
        elif sysAperData[0] == 5: # Object Cone Angle
            print(("WARNING: Scaling OCA aperture type may be incorrect for {lF}"
                   .format(lF=lensFile)))
            ret = 1
        #Get the number of surfaces
        numSurf = 0
        recSystemData_g = self.zGetSystem() #Get the current system parameters
        numSurf = recSystemData_g[0]
        #print("Number of surfaces in the lens: ", numSurf)
        if recSystemData_g[4] > 0:
            print("Warning: Ray aiming is ON in {lF}. But cannot scale"
                  " Pupil Shift values.".format(lF=lensFile))

        #Scale individual surface properties in the LDE
        for surfNum in range(0,numSurf+1): #Start from the object surface ... to scale thickness if not infinity
            #Scale the basic data common to all surface types such as radius, thickness
            #and semi-diameter
            surfName = self.zGetSurfaceData(surfNum,0)
            curv = self.zGetSurfaceData(surfNum,2)
            thickness = self.zGetSurfaceData(surfNum,3)
            semiDiam = self.zGetSurfaceData(surfNum,5)
            ##print("Surf#:",surfNum,"Name:",surfName,"Curvature:",curv,"Thickness:",thickness,"Semi-Diameter:", semiDiam)
            #scale the basic data
            scaledCurv = self.zSetSurfaceData(surfNum,2,curv/factor)
            if thickness < 1.0E+10: #Scale the thickness if it not Infinity (-1.0E+10 in Zemax)
                scaledThickness = self.zSetSurfaceData(surfNum,3,factor*thickness)
            scaledSemiDiam = self.zSetSurfaceData(surfNum,5,factor*semiDiam)
            ##print("scaled", surfNum,surfName,scaledCurv,scaledThickness,scaledSemiDiam)

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
                    for i in range(3,binSurMaxNum[surfName]): # scaling of terms 3 to 232, p^480
                                                              # for Binary1 and Binary 2 respectively
                        if i > numBTerms + 2: #(+2 because the terms starts from par 3)
                            break
                        else:
                            epar = self.zGetExtra(surfNum,i)
                            epar_ret = self.zSetExtra(surfNum,i,factor*epar)
            elif surfName == 'BINARY_3':
                #Scaling of parameters in the LDE
                par1 = self.zGetSurfaceParameter(surfNum,1) # R2
                par1_ret = self.zSetSurfaceParameter(surfNum,1,factor*par1)
                par4 = self.zGetSurfaceParameter(surfNum,4) # A2, need to scale A2 before A1,
                                                            # because A2>A1>0.0 always
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
                print(("WARNING: Scaling for surf type {sN} in file {lF} not implemented!!"
                      .format(sN=surfName,lF=lensFile)))
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
                tField = list(fieldDataTuple[i])
                tField[0],tField[1] = factor*tField[0],factor*tField[1]
                fieldDataTupleScaled.append(tuple(tField))
            fieldDataTupleScaled = self.zSetFieldTuple(fType,fNorm,
                                                 tuple(fieldDataTupleScaled))
        return ret


    def calculateHiatus(self,txtFileName2Use=None,keepFile=False):
        """Calculate the Hiatus.

        The hiatus, also known as the Null space, or nodal space, or the interstitium
        is the distance between the two principal planes.

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

        """
        if txtFileName2Use != None:
            textFileName = txtFileName2Use
        else:
            #cd = os.getcwd()
            cd = os.path.dirname(os.path.realpath(__file__))
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
        #We are creating a list of lines by purpose, see note 2 (decisions for this fn)
        line_list = fileref.readlines()
        fileref.close()

        for line_num,line in enumerate(line_list):
            #Extract the image surface distance from the global ref sur (surface 1)
            sectionString = ("GLOBAL VERTEX COORDINATES, ORIENTATIONS,"
                             " AND ROTATION/OFFSET MATRICES:")
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
# Helper functions to process data from ZEMAX DDE server. This is especially
# convenient for processing replies from Zemax for those function calls that
# outputs exactly same reply structure.
# ***************************************************************************

def regressLiteralType(x):
    """The function returns the literal with its proper type, such as int, float,
    or string from the input string x.
    Example:
        regressLiteralType("1")->1
        regressLiteralType("1.0")->1.0
        regressLiteralType("1e-3")->0.001
        regressLiteralType("YUV")->'YUV'
        regressLiteralType("YO8")->'YO8'
    """
    try:
        float(x)  # Test for numeric or string
        lit = float(x) if set(['.','e','E']).intersection(x) else int(x)
    except ValueError:
        lit = str(x)
    return lit

def process_get_set_NSCProperty(code,reply):
    """Process reply for functions zGetNSCProperty and zSETNSCProperty"""
    rs = reply.rstrip()
    if rs == 'BAD COMMAND':
        nscPropData = -1
    else:
        if code in (0,1,4,5,6,11,12,14,18,19,27,28,84,86,92,117,123):
            nscPropData = str(rs)
        elif code in (2,3,7,9,13,15,16,17,20,29,81,91,101,102,110,111,113,
                      121,141,142,151,152,153161,162,171,172,173):
            nscPropData = int(float(rs))
        else:
            nscPropData = float(rs)
    return nscPropData

def process_get_set_Operand(column, reply):
    """Process reply for functions zGetOperand and zSetOperand"""
    rs = reply.rstrip()
    if column == 1:
        #ensure that it is a string ... as it is supposed to return the operand
        if isinstance(regressLiteralType(rs),str):
            return str(rs)
        else:
            return -1
    elif column in (2,3):
        return int(float(rs))
    else:
        return float(rs)

def process_get_set_Solve(reply):
    """Process reply for functions zGetSolve and zSetSolve"""
    reply = reply.rstrip()
    rs = reply.split(",")
    if 'BAD COMMAND' in rs:
        return -1
    else:
        return tuple([regressLiteralType(x) for x in rs])

def process_get_set_SystemProperty(code,reply):
    """Process reply for functions zGetSystemProperty and zSetSystemProperty"""
    #Convert reply to proper type
    if code in  (102,103, 104,105,106,107,108,109,110): # unexpected cases
        sysPropData = reply
    elif code in (16,17,23,40,41,42,43): # string
        sysPropData = str(reply)
    elif code in (11,13,24,53,54,55,56,60,61,62,63,71,72,73,77,78): # floats
        sysPropData = float(reply)
    else:
        sysPropData = int(float(reply))      # integer
    return sysPropData

def process_get_set_Tol(operandNumber,reply):
    """Process reply for functions zGetTol and zSetTol"""
    rs = reply.rsplit(",")
    tolType = [rs[0]]
    tolParam = [float(e) if i in (2,3) else int(float(e))
                                 for i,e in enumerate(rs[1:])]
    toleranceData = tuple(tolType + tolParam)
    return toleranceData


# ***************************************************************************
#                      TEST (SOME OF) THE FUNCTIONS
# ***************************************************************************
# Initial code testing used to be done using the following lines of code.
# Currently all functionality are being tested using the unit test module.
# The test_PyZDDE() function are left for quick test. The test_PyZDDE() function
# will not be executed if the module is imported! In order to execute the
# test_PyZDDE() function, execute this (pyzdde.py) module. It may prove to be
#useful to quickly test your system.

def test_PyZDDE():
    """Test the pyzdde module functions"""
    zmxfp = os.path.dirname(os.path.realpath(__file__))+'\\ZMXFILES\\'
    # Create PyZDDE object(s)
    link0 = PyZDDE()
    link1 = PyZDDE()
    link2 = PyZDDE()  # this object shall be deleted randomly

    print("\nTEST: zDDEInit()")
    print("---------------")
    status = link0.zDDEInit()
    print("Status for link 0:", status)
    assert status == 0
    print("App Name for Link 0:", link0.appName)
    print("Connection status for Link 0:", link0.connection)
    time.sleep(0.1)   # Not required, but just for observation

    #link1 = PyZDDE()
    status = link1.zDDEInit()
    print("Status for link 1:",status)
    #assert status == 0   # In older versions of Zemax, unable to create second
                          # communication link.
    if status != 0:
        warnings.warn("Couldn't create second channel.\n")
    else:
        print("App Name for Link 1:", link1.appName)
        print("Connection status for Link 1:", link1.connection)
    time.sleep(0.1)   # Not required, but just for observation

    print("\nTEST: zGetDate()")
    print("----------------")
    print("Date: ", link0.zGetDate().rstrip())  # strip off the newline char

    print("\nTEST: zGetSerial()")
    print("------------------")
    ser = link0.zGetSerial()
    print("Serial #:", ser)

    print("\nTEST: zGetVersion()")
    print("----------------")
    print("version number: ", link0.zGetVersion())

    print("\nTEST: zSetTimeout()")
    print("------------------")
    link0.zSetTimeout(3)

    #Delete link2 randomly
    print("\nTEST: Random deletion of object")
    print("--------------------------------")
    print("Deleting object link2")
    del link2

    print("\nTEST: zLoadFile()")
    print("-------------------")
    filename = zmxfp+"nonExistantFile.zmx"
    ret = link0.zLoadFile(filename)
    assert ret == -999
    filename = zmxfp+"Cooke 40 degree field.zmx"
    ret = link0.zLoadFile(filename)
    assert ret == 0
    print("zLoadFile test successful")

    if link1.connection:
        print("\nTEST: zLoadFile() @ link 1 (second channel)")
        print("-------------------")
        filename = zmxfp+"Double Gauss 5 degree field.zmx"
        assert ret == 0
        print("zLoadFile test @ link 1 successful")

    print("\nTEST: zPushLensPermission()")
    print("---------------------------")
    status = link0.zPushLensPermission()
    if status:
        print("Extensions are allowed to push lens.")

        print("\nTEST: zPushLens()")
        print("-----------------")
        # First try to push a lens with invalid flag argument
        try:
            ret = link0.zPushLens(updateFlag=10)
        except:
            info = sys.exc_info()
            print("Exception error:", info[0])
            #assert info[0] == 'exceptions.ValueError'
            assert cmp(str(info[0]),"<type 'exceptions.ValueError'>") == 0

        # TEST ALL FUNCTIONS THAT REQUIRE PUSHLENS() ... HERE!
        #Push lens without any parameters
        ret = link0.zPushLens()
        if ret ==0:
            print("Lens update without any arguments suceeded. ret value = ", ret)
        else:
            print("Lens update without any arguments FAILED. ret value = ", ret)
        #Push lens with some valid parameters
        ret = link0.zPushLens(updateFlag=1)
        if ret == 0:
            print("Lens update with flag=1 suceeded. ret value = ", ret)
        else:
            print("Lens update with flag=1 FAILED. ret value = ", ret)

    else: # client do not have permission to push lens
        print("Extensions are not allowed to push lens. Please enable it.")

    #Continue with other tests
    print("\nTEST: zGetTrace()")
    print("------------------")
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
    print("zGetTrace test successful")

    print("\nTEST: zGetRefresh()")
    print("------------------")
    status = link0.zGetRefresh()
    if status == 0:
        print("Refresh successful")
    else:
        print("Refresh FAILED")

    print("\nTEST: zSetSystem()")
    print("-----------------")
    unitCode,stopSurf,rayAimingType = 0,4,0  # mm, 4th,off
    useEnvData,temp,pressure,globalRefSurf = 0,20,1,1 # off, 20C,1ATM,ref=1st surf
    systemData_s = link0.zSetSystem(unitCode,stopSurf,rayAimingType,useEnvData,
                                              temp,pressure,globalRefSurf)
    print(systemData_s)

    print("\nTEST: zGetSystem()")
    print("-----------------")
    systemData_g = link0.zGetSystem()
    print(systemData_g)

    assert systemData_s == systemData_g
    print("zSetSystem() and zGetSystem() test successful")

    print("\nTEST: zGetPupil()")
    print("------------------")
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

    # Start a basic design with a new lens
    print("\nTEST: zNewLens()")
    print("----------------")
    retVal = link0.zNewLens()
    assert retVal == 0
    print("zNewLens() test successful")

    #Set (new) system parameters:
    #Get the current stop position (it should be 1, as it is a new lens)
    sysPara = link0.zGetSystem()
    # set unitCode (mm), stop-surface, ray-aiming, ... , global surface reference
    sysParaNew = link0.zSetSystem(0,sysPara[2],0,0,20,1,-1) # Set the image plane as Global ref surface

    print("\nTEST: zSetSystemAper():")
    print("-------------------")
    systemAperData_s = link0.zSetSystemAper(0,sysPara[2],25) # sysAper = 25 mm, EPD
    assert systemAperData_s[0] == 0  # Confirm aperType = EPD
    assert systemAperData_s[1] == sysPara[2]  # confirm stop surface number
    assert systemAperData_s[2] == 25  # confirm EPD value is 25 mm
    print("zSetSystemAper() test successful")

    print("\nTEST: zGetSystemAper():")
    print("-----------------------")
    systemAperData_g = link0.zGetSystemAper()
    assert systemAperData_s == systemAperData_g
    print("zGetSystemAper() test successful")

    print("\nTEST: zInsertSurface()")
    print("--------------------")
    retVal = link0.zInsertSurface(1)
    assert retVal == 0
    print("zInsertSurface() successful")

    print("\nTEST: zSetAperture()")
    print("---------------------")
    #aptInfo = link0.zSetAperture()
    pass
    #ToDo

    print("\nTEST: zGetAperture()")
    print("---------------------")
    #aptInfo = link0.zGetAperture()
    pass
    #ToDo

    print("\nTEST: zSetField()")
    print("---------------------")
    fieldData = link0.zSetField(0,0,2) # type = angle; 2 fields; rect normalization (default)
    print("fieldData: ",fieldData)
    assert fieldData[0]==0; assert fieldData[1]==2;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zSetField(0,0,3,1)
    print("fieldData: ",fieldData)
    assert fieldData[0]==0; assert fieldData[1]==3;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zSetField(1,0,0) # 1st field, on-axis x, on-axis y, weight = 1 (default)
    print("fieldData: ",fieldData)
    assert fieldData==(0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    fieldData = link0.zSetField(2,0,5,2.0,0.5,0.5,0.5,0.5,0.5)
    print("fieldData: ",fieldData)
    assert fieldData==(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5)
    fieldData = link0.zSetField(3,0,10,1.0,0.0,0.0,0.0)
    print("fieldData: ",fieldData)
    assert fieldData==(0.0, 10.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    print("zSetField() test successful")

    print("\nTEST: zGetField()")
    print("---------------------")
    fieldData = link0.zGetField(0)
    print("fieldData: ",fieldData)
    assert fieldData[0]==0; assert fieldData[1]==3;
    #assert fieldData[4]== 1; (normalization)
    fieldData = link0.zGetField(2)
    print("fieldData: ",fieldData)
    assert fieldData==(0.0, 5.0, 2.0, 0.5, 0.5, 0.5, 0.5, 0.5)
    print("zGetField() successful")

    print("\nTEST: zSetFieldTuple()")
    print("---------------------")
    iFieldDataTuple = ((0.0,0.0,1.0,0.0,0.0,0.0,0.0,0.0), # field1: xf=0.0,yf=0.0,wgt=1.0,
                                                          # vdx=vdy=vcx=vcy=van=0.0
                       (0.0,5.0,1.0),                     # field2: xf=0.0,yf=5.0,wgt=1.0
                       (0.0,10.0))                        # field3: xf=0.0,yf=10.0
    oFieldDataTuple = link0.zSetFieldTuple(0,1,iFieldDataTuple)
    for i in range(len(iFieldDataTuple)):
        print("oFieldDataTuple, field {} : {}".format(i,oFieldDataTuple[i]))
        assert oFieldDataTuple[i][:len(iFieldDataTuple[i])]==iFieldDataTuple[i]
    print("zSetFieldTuple() test successful")

    print("\nTEST: zGetFieldTuple()")
    print("----------------------")
    fieldDataTuple = link0.zGetFieldTuple()
    assert fieldDataTuple==oFieldDataTuple
    print("zGetFieldTuple() test successful")

    print("\nTEST: zSetWave()")
    print("-----------------")
    wavelength1 = 0.48613270
    wavelength2 = 0.58756180
    waveData = link0.zSetWave(0,1,2)
    print("Primary wavelength number = ", waveData[0])
    print("Total number of wavelengths set = ",waveData[1])
    assert waveData[0]==1; assert waveData[1]==2
    waveData = link0.zSetWave(1,wavelength1,0.5)
    print("Wavelength 1: ",waveData[0])
    assert waveData[0]==wavelength1;assert waveData[1]==0.5
    waveData = link0.zSetWave(2,wavelength2,0.5)
    print("Wavelength 2: ",waveData[0])
    assert waveData[0]==wavelength2;assert waveData[1]==0.5
    print("zSetWave test successful")

    print("\nTEST: zGetWave()")
    print("-----------------")
    waveData = link0.zGetWave(0)
    assert waveData[0]==1;assert waveData[1]==2
    print(waveData)
    waveData = link0.zGetWave(1)
    assert waveData[0]==wavelength1;assert waveData[1]==0.5
    print(waveData)
    waveData = link0.zGetWave(2)
    assert waveData[0]==wavelength2;assert waveData[1]==0.5
    print(waveData)
    print("zGetWave test successful")

    print("\nTEST:zSetWaveTuple()")
    print("-------------------------")
    wavelengths = (0.48613270,0.58756180,0.65627250)
    weights = (1.0,1.0,1.0)
    iWaveDataTuple = (wavelengths,weights)
    oWaveDataTuple = link0.zSetWaveTuple(iWaveDataTuple)
    print("Output wave data tuple",oWaveDataTuple)
    assert oWaveDataTuple==iWaveDataTuple
    print("zSetWaveTuple() test successful")

    print("\nTEST:zGetWaveTuple()")
    print("-------------------------")
    waveData = link0.zGetWaveTuple()
    print("Wave data tuple =",waveData)
    assert oWaveDataTuple==waveData
    print("zGetWaveTuple() test successful")

    print("\nTEST: zSetPrimaryWave()")
    print("-----------------------")
    primaryWaveNumber = 2
    waveData = link0.zSetPrimaryWave(primaryWaveNumber)
    print("Primary wavelength number =", waveData[0])
    print("Total number of wavelengths =", waveData[1])
    assert waveData[0]==primaryWaveNumber
    assert waveData[1]==len(wavelengths)
    print("zSetPrimaryWave() test successful")

    print("\nTEST: zQuickFocus()")
    print("---------------------")
    retVal = link0.zQuickFocus()
    print("zQuickFocus() test retVal = ", retVal)
    assert retVal == 0
    print("zQuickFocus() test successful")

    # Finished all tests. Perform the last test and done!
    print("\nTEST: zDDEClose()")
    print("----------------")
    status = link0.zDDEClose()
    print("Communication link 0 with ZEMAX terminated")
    status = link1.zDDEClose()
    print("Communication link 1 with ZEMAX terminated")


if __name__ == '__main__':
    import os, time
    test_PyZDDE()
