#-------------------------------------------------------------------------------
# Name:        zdde.py
# Purpose:     Python based DDE link with ZEMAX server, similar to Matlab based
#              MZDDE toolbox.
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2014
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.7.6
#-------------------------------------------------------------------------------
"""PyZDDE, which is a toolbox written in Python, is used for communicating with
ZEMAX using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol.
The docstring examples in the functions assume that PyZDDE is imported as
``import pyzdde.zdde as pyz`` and a PyZDDE communication object is then created
as ``ln = pyz.createLink()`` or ``ln = pyz.PyZDDE(); ln.zDDEInit()``.
"""
from __future__ import division
from __future__ import print_function
import sys as _sys
import os as _os
import collections as _co
import subprocess as _subprocess
import math as _math
import time as _time
import datetime as _datetime
import re as _re
import shutil as _shutil
# import warnings as _warnings
import codecs as _codecs
# Try to import IPython if it is available (for notebook helper functions)
try:
    from IPython.core.display import display as _display
    from IPython.core.display import Image as _Image
except ImportError:
    _global_IPLoad = False
else:
    _global_IPLoad = True

# Determine if in IPython Environment
try: # get_ipython() method is not available in IPython versions prior to 2.0
    from IPython import get_ipython as _get_ipython
except:
    _global_in_IPython_env = False
else:
    if _get_ipython(): # if global interactive shell instance is available
        _global_in_IPython_env = True
    else:
        _global_in_IPython_env = False
# Try to import Matplotlib's imread
try:
    import matplotlib.image as _matimg
except ImportError:
    _global_mpl_img_load = False
else:
    _global_mpl_img_load = True

# Import intra-package modules
# TODO!!! Appending current dir trick should be removed once packaging is used
_currDir = _os.path.dirname(_os.path.realpath(__file__))
_index = _currDir.find('pyzdde')
_pDir = _currDir[0:_index-1]
if _pDir not in _sys.path:
    _sys.path.append(_pDir)

# The first module to import that is not one of the standard modules MUST
# be the config module as it sets up the different global and settings variables
import pyzdde.config as _config
_global_pyver3 = _config._global_pyver3
_global_use_unicode_text = _config._global_use_unicode_text

# DDEML communication module
import pyzdde.ddeclient as _dde

if _global_pyver3:
   _izip = zip
   _imap = map
else:
    from itertools import izip as _izip, imap as _imap

import pyzdde.zcodes.zemaxbuttons as zb
import pyzdde.zcodes.zemaxoperands as zo
import pyzdde.utils.pyzddeutils as _putils

# Constants
_DEBUG_PRINT_LEVEL = 0 # 0=No debug prints, but allow all essential prints
                       # 1 to 2 levels of debug print, 2 = print all

_MAX_PARALLEL_CONV = 2  # Max no of simul. conversations possible with Zemax
_system_aperture = {0 : 'EPD',
                    1 : 'Image space F/#',
                    2 : 'Object space NA',
                    3 : 'Float by stop',
                    4 : 'Paraxial working F/#',
                    5 : 'Object cone angle'}

# Helper function for debugging
def _debugPrint(level, msg):
    """Internal helper function to print debug messages

    Parameters
    ----------
    level : integer (0, 1 or 2)
        0 = message will definitely be printed;
        1 or 2 = message will be printed if ``level >= _DEBUG_PRINT_LEVEL``.
    msg : string
        message to print
    """
    global _DEBUG_PRINT_LEVEL
    if level <= _DEBUG_PRINT_LEVEL:
        print("DEBUG PRINT, module - zdde (Level " + str(level)+ "): " + msg)

# ***************
# Module methods
# ***************
# bind functions from utils module
cropImgBorders = _putils.cropImgBorders
imshow = _putils.imshow
# bind functions from zemax buttons module
findZButtonCode = zb.findZButtonCode
getZButtonCount = zb.getZButtonCount
isZButtonCode   = zb.isZButtonCode
showZButtonList = zb.showZButtonList
showZButtonDescription = zb.showZButtonDescription
# bind functions from zemax operand module
findZOperand = zo.findZOperand
getZOperandCount = zo.getZOperandCount
isZOperand = zo.isZOperand
showZOperandList = zo.showZOperandList
showZOperandDescription = zo.showZOperandDescription

_global_dde_linkObj = {}

def createLink():
    """Create a DDE communication link with Zemax

    Usage: ``import pyzdde.zdde as pyz; ln = pyz.createLink()``

    Helper function to create, initialize and return a PyZDDE communication
    object.

    Parameters
    ----------
    None

    Returns
    -------
    link : object
        a PyZDDE communication object if successful, else ``None``.

    Notes
    -----
    1. This module level method may used instead of \
    ``ln = pyz.PyZDDE(); ln.zDDEInit()``.
    2. Zemax application must be running.

    See Also
    --------
    closeLink(), zDDEInit()
    """
    global _global_dde_linkObj
    global _MAX_PARALLEL_CONV
    dlen = len(_global_dde_linkObj)
    if dlen < _MAX_PARALLEL_CONV:
        link = PyZDDE()
        status = link.zDDEInit()
        if not status:
            _global_dde_linkObj[link] = link.appName  # This can be something more useful later
            _debugPrint(1,"Link created. Link Dict = {}".format(_global_dde_linkObj))
            return link
        else:
            print("Could not initiate instance.")
            return None
    else:
        print("Link not created. Reached maximum allowable live link of ",
              _MAX_PARALLEL_CONV)
        return None

def closeLink(link=None):
    """Close DDE communication link with Zemax

    Usage: ``pyz.closeLink([ln])``

    Helper function, for closing DDE communication.

    Parameters
    ----------
    link : PyZDDE link object, optional
        If a specific link object is not given, all existing links are closed.

    Returns
    -------
    None

    See Also
    --------
    zDDEClose() :
        PyZDDE instance method to close a link.
        Use this method (as ``ln.zDDEClose()``) if the link was created as \
        ``ln = pyz.PyZDDE(); ln.zDDEInit()``
    close() :
        Another instance method to close a link for easy typing.
        Use this method (as ``ln.close()``) or ``pyz.closeLink(ln)`` if the \
        link was created as ``ln = pyz.createLink()``
    """
    global _global_dde_linkObj
    dde_closedLinkObj = []
    if link:
        link.zDDEClose()
        dde_closedLinkObj.append(link)
    else:
        for link in _global_dde_linkObj:
            link.zDDEClose()
            dde_closedLinkObj.append(link)
    for item in dde_closedLinkObj:
        _global_dde_linkObj.pop(item)

def setTextEncoding(txt_encoding=0):
    """Sets PyZDDE text encoding to match the TXT encoding in Zemax

    Usage: ``pyz.setTextEncoding([txt_encoding])``

    Parameters
    ----------
    txt_encoding : integer (0 or 1)
        0 = ASCII; 1 = UNICODE

    Returns
    -------
    status : string
        current setting (& information if the setting was changed or not)

    Notes
    -----
    Not required to set the encoding for every new session as PyZDDE stores
    the setting.

    See Also
    --------
    getTextEncoding()
    """
    global _global_use_unicode_text
    if _global_use_unicode_text and txt_encoding:
        print('TXT encoding is UNICODE; no change required')
    elif not _global_use_unicode_text and not txt_encoding:
        print('TXT encoding is ASCII; no change required')
    elif not _global_use_unicode_text and txt_encoding:
        if _config.setTextEncoding(txt_encoding=1):
            _global_use_unicode_text = True
            print('Successfully changed to UNICODE')
        else:
            print("ERROR: Couldn't change settings")
    elif _global_use_unicode_text and not txt_encoding:
        if _config.setTextEncoding(txt_encoding=0):
            _global_use_unicode_text = False
            print('Successfully changed to ASCII')
        else:
            print("ERROR: Couldn't change settings")

def getTextEncoding():
    """Returns the current text encoding set in PyZDDE

    Usage: ``pyz.getTextEncoding()``

    Parameters
    ----------
    None

    Returns
    -------
    encoding : string
        'ascii' or 'unicode'

    See Also
    --------
    setTextEncoding
    """
    return _config.getTextEncoding()

# PyZDDE class' utility function (for internal use)
def _createAppNameDict(maxElements):
    """Internal function to create a pool (dictionary) of possible app-names

    Parameters
    ----------
    maxElements : integer
        maximum elements in the dictionary

    Returns
    -------
    appNameDict : dictionary
        dictionary of app-names (keys) with values, set to False, indicating
        name hasn't been taken.
    """
    appNameDict = {}
    appNameDict['ZEMAX'] = False
    for i in range(1, maxElements):
        appNameDict['ZEMAX'+str(i)] = False
    return appNameDict

def _getAppName(appNameDict):
    """Return available name from the pool of app-names.
    """
    if not appNameDict['ZEMAX']:
        appNameDict['ZEMAX'] = True
        return 'ZEMAX'
    else:
        k_available = None
        for k, v in appNameDict.items():
            if not v:
                k_available = k
                break
        if k_available:
            appNameDict[k_available] = True
            return k_available
        else:
            return None


# ******************
# PyZDDE class
# ******************
class PyZDDE(object):
    """PyZDDE class for communicating with Zemax

    There are two ways of instantiating and initiating a PyZDDE object:

    1. Instantiate using ``ln = pyz.PyZDDE()`` and then initiate \
    using ``ln.zDDEInit()`` or
    2. ``pyz.createLink()`` instantiates & initiates a PyZDDE object & \
    returns (recommended way)
    """
    __chNum =  0  # channel Number; there is no restriction on number of ch
    __liveCh = 0  # no. of live/ simul channels; Can't be > _MAX_PARALLEL_CONV
    __server = 0
    __appNameDict = _createAppNameDict(_MAX_PARALLEL_CONV)  # {'ZEMAX': False, 'ZEMAX1': False}

    def __init__(self):
        """Creates an instance of PyZDDE class

        Usage: ``ln = pyz.PyZDDE()``

        Parameters
        ----------
        None

        Returns
        -------
        ln : PyZDDE object

        Notes
        -----
        1. Following the creation of PyZDDE object, initiate the communication
        channel as ``ln.zDDEInit()``
        2. Consider using the module level function ``pyz.createLink()`` to create
        and initiate a DDE channel instead of ``ln = pyz.PyZDDE(); ln.zDDEInit()``

        See Also
        --------
        createLink()
        """
        PyZDDE.__chNum += 1   # increment channel count
        self.appName = _getAppName(PyZDDE.__appNameDict) or '' # wicked :-)
        self.appNum = PyZDDE.__chNum # unique & immutable identity of each instance
        self.connection = False  # 1/0 depending on successful connection or not
        self.macroPath = None    # variable to store macro path

    def __repr__(self):
        return ("PyZDDE(appName=%r, appNum=%r, connection=%r, macroPath=%r)" %
                (self.appName, self.appNum, self.connection, self.macroPath))

    def __hash__(self):
        # for storing in internal dictionary
        return hash(self.appNum)

    def __eq__(self, other):
        return (self.appNum == other.appNum)

    # ZEMAX <--> PyZDDE client connection methods
    #--------------------------------------------
    def zDDEInit(self):
        """Initiates DDE link with Zemax server.

        Usage: ``ln.zDDEInit()``

        Parameters
        ----------
        None

        Returns
        -------
        status : integer (0 or -1)
            0 = DDE Zemax link successful;
            -1 = DDE link couldn't be established.

        See Also
        --------
        createLink(), zDDEClose(), zDDEStart(), zSetTimeout()
        """
        _debugPrint(1,"appName = " + self.appName)
        _debugPrint(1,"liveCh = " + str(PyZDDE.__liveCh))
        # do this only one time or when there is no channel
        if PyZDDE.__liveCh==0:
            try:
                PyZDDE.__server = _dde.CreateServer()
                PyZDDE.__server.Create("ZCLIENT")  # Name of the client
                _debugPrint(2, "PyZDDE.__server = " + str(PyZDDE.__server))
            except Exception as err1:
                _sys.stderr.write("{err}: Another application may be"
                                 " using a DDE server!".format(err=str(err1)))
                return -1
        # Try to create individual conversations for each ZEMAX application.
        self.conversation = _dde.CreateConversation(PyZDDE.__server)
        _debugPrint(2, "PyZDDE.converstation = " + str(self.conversation))
        try:
            self.conversation.ConnectTo(self.appName," ")
        except Exception as err2:
            _debugPrint(2, "Exception occured at attempt to call ConnecTo."
                        " Error = {err}".format(err=str(err2)))
            if self.__liveCh >= _MAX_PARALLEL_CONV:
                _sys.stderr.write("ERROR: {err}. \nMore than {liveConv} "
                "simultaneous conversations not allowed!\n"
                .format(err=str(err2), liveConv =_MAX_PARALLEL_CONV))
            else:
                _sys.stderr.write("ERROR: {err}.\nZEMAX may not be running!\n"
                                 .format(err=str(err2)))
            # should close the DDE server if it exist
            self.zDDEClose()
            _debugPrint(2,"PyZDDE server: " + str(PyZDDE.__server))
            return -1
        else:
            _debugPrint(1,"Zemax instance successfully connected")
            PyZDDE.__liveCh += 1 # increment the number of live channels
            self.connection = True
            return 0

    def close(self):
        """Helper function to close current communication link

        Usage: ``ln.close()``

        Parameters
        ----------
        None

        Returns
        -------
        None

        Notes
        -----
        This bounded method provides a quick alternative way to close link rather
        than calling the module function ``pyz.closeLink()``.

        See Also
        --------
        zDDEClose() :
            PyZDDE instance method to close a link.
            Use this method (as ``ln.zDDEClose()``) if the link was created as \
            ``ln = pyz.PyZDDE(); ln.zDDEInit()``
        closeLink() :
            A moudle level function to close a link.
            Use this method (as ``pyz.closeLink(ln)``) or ``ln.close()`` if the \
            link was created as ``ln = pyz.createLink()``
        """
        closeLink(self)

    def zDDEClose(self):
        """Close the DDE link with Zemax server.

        Usage: ``ln.zDDEClose()``

        Parameters
        ----------
        None

        Returns
        -------
        status : integer
            0 on success.

        Notes
        -----
        Use this bounded method to close link if the link was created using the
        idiom ``ln = pyz.PyZDDE(); ln.zDDEInit()``. If however, the link was
        created using ``ln = pyz.createLink()``, use either ``pyz.closeLink()``
        or ``ln.close()``.
        """
        if PyZDDE.__server and not PyZDDE.__liveCh:
            PyZDDE.__server.Shutdown(self.conversation)
            PyZDDE.__server = 0
            _debugPrint(2,"server shutdown as ZEMAX is not running!")
        elif PyZDDE.__server and self.connection and PyZDDE.__liveCh == 1:
            PyZDDE.__server.Shutdown(self.conversation)
            self.connection = False
            PyZDDE.__appNameDict[self.appName] = False # make the name available
            self.appName = ''
            PyZDDE.__liveCh -= 1  # This will become zero now. (reset)
            PyZDDE.__server = 0   # previous server obj should be garbage collected
            _debugPrint(2,"server shutdown")
        elif self.connection:  # if additional channels were successfully created.
            PyZDDE.__server.Shutdown(self.conversation)
            self.connection = False
            PyZDDE.__appNameDict[self.appName] = False # make the name available
            self.appName = ''
            PyZDDE.__liveCh -= 1
            _debugPrint(2,"liveCh decremented without shutting down DDE channel")
        else:   # if zDDEClose is called by an object which didn't have a channel
            _debugPrint(2, "Nothing to do")
        return 0

    def zSetTimeout(self, time):
        """Set global timeout value, in seconds, for all Zemax DDE calls.

        Parameters
        ----------
        time: integer
            timeout value in seconds (if float is given, it is rounded to integer)

        Returns
        -------
        timeout : integer
            the set timeout value in seconds

        Notes
        -----
        This is a global timeout value. Some methods provide means to set
        individual timeout values.

        See Also
        --------
        zDDEInit()
        """
        self.conversation.SetDDETimeout(round(time))
        return self.conversation.GetDDETimeout()


    def zGetTimeout(self):
        """Returns the current value of the global timeout in seconds

        Parameters
        ----------
        None

        Returns
        -------
        timeout : integer
            globally set timeout value in seconds
        """
        return self.conversation.GetDDETimeout()

    def _sendDDEcommand(self, cmd, timeout=None):
        """Method to send command to DDE client
        """
        global _global_pyver3
        reply = self.conversation.Request(cmd, timeout)
        if _global_pyver3:
            reply = reply.decode('ascii').rstrip()
        return reply

    def __del__(self):
        """Destructor"""
        _debugPrint(2,"Destructor called")
        self.zDDEClose()

    # ZEMAX control/query methods
    #----------------------------
    def zCloseUDOData(self, bufferCode):
        """Close the User Defined Operand buffer allowing optimizer to proceed.

        Parameters
        ----------
        bufferCode : integer
            buffercode is an integer value provided by Zemax to the client that
            uniquely identifies the correct lens.

        Returns
        -------
        retVal : ?

        See Also
        --------
         zGetUDOSystem(), zSetUDOItem()
        """
        return int(self._sendDDEcommand("CloseUDOData,{:d}".format(bufferCode)))

    def zDeleteConfig(self, number):
        """Deletes an existing configuration (column) in the multi-configuration
        editor

        Parameters
        ----------
        number : integer
            configuration number to delete

        Returns
        -------
        deleted_config_num : integer
            configuration number deleted

        Notes
        -----
        After deleting the configuration, all succeeding configurations are
        re-numbered.

        See Also
        --------
        zInsertConfig()

        zDeleteMCO() :
            (TIP) use zDeleteMCO to delete a row/operand
        """
        return int(self._sendDDEcommand("DeleteConfig,{:d}".format(number)))

    def zDeleteMCO(self, operandNumber):
        """Deletes an existing operand (row) in the multi-configuration editor.

        Parameters
        ----------
        operandNumber : integer
            operand number (row in the MCE) to delete

        Returns
        -------
        newNumberOfOperands : integer
            new number of operands

        Notes
        -----
        After deleting the row, all succeeding rows (operands) are re-numbered.

        See Also
        --------
        zInsertMCO()
        zDeleteConfig() :
            (TIP) Use zDeleteConfig() to delete a column/configuration.
        """
        return int(self._sendDDEcommand("DeleteMCO,"+str(operandNumber)))

    def zDeleteMFO(self, operand):
        """Deletes an optimization operand (row) in the merit function editor

        Parameters
        ----------
        operand : integer
            Operand number (- 1 <= operand <= number_of_operands)

        Returns
        -------
        newNumOfOperands : integer
            the new number of operands

        See Also
        --------
        zInsertMFO()
        """
        return int(self._sendDDEcommand("DeleteMFO,{:d}".format(operand)))

    def zDeleteObject(self, surfaceNumber, objectNumber):
        """Deletes the NSC object identified by the ``objectNumber`` and the
        surface identified by ``surfaceNumber``

        Parameters
        ----------
        surfaceNumber : integer
            surface number of Non-Sequential Component surface
        objectNumber : integer
            object number in the NSC editor

        Returns
        -------
        status : integer (0 or -1)
            0 if successful, -1 if it failed

        Notes
        -----
        1. (from MZDDE) The ``surfaceNumber`` is 1 if the lens is purely NSC mode.
        2. If no more objects are present it simply returns 0.

        See Also
        --------
        zInsertObject()
        """
        cmd = "DeleteObject,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            return -1
        else:
            return int(float(rs))

    def zDeleteSurface(self, surfaceNumber):
        """Deletes an existing surface identified by ``surfaceNumber``

        Parameters
        ----------
        surfaceNumber : integer
            the surface number of the surface to be deleted

        Returns
        -------
        status : integer
            0 if successful

        .. warning:: Although you cannot delete an object the function \
        doesn't return an error (returns 0 instead).

        See Also
        --------
        zInsertSurface()
        """
        cmd = "DeleteSurface,{:d}".format(surfaceNumber)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply))

    def zExecuteZPLMacro(self, zplMacroCode, timeout=None):
        """Executes a ZPL macro present in the <data>/Macros folder.

        Parameters
        ----------
        zplMacroCode : string
            The first 3 letters (case-sensitive) of the ZPL macro present
            in the <data>/Macros folder
        timeout : integer
            timeout value in seconds

        Returns
        --------
        status : integer (0 or 1)
            0 = successfully executed the ZPL macro;
            -1 = macro code is incorrect & error code returned by Zemax

        Notes
        -----
        If the macro path is different from the default macro path at \
        ``<data>/Macros``, then first use ``zSetMacroPath()`` to set the \
        macropath and then use ``zExecuteZPLMacro()``.

        .. warning::

          1. can only "execute" an existing ZPL macro. i.e. you can't \
             create a ZPL macro on-the-fly and execute it.
          2. If it is required to redirect the result of executing the ZPL \
             to a text file, modify the ZPL macro in the following way:

            -   Add the following two lines at the beginning of the file:
                ``CLOSEWINDOW`` # to suppress the display of default text window
                ``OUTPUT "full_path_with_extension_of_result_fileName"``
            -   Add the following line at the end of the file:
                ``OUTPUT SCREEN`` # close the file and re-enable screen printing

          3. If there are more than two macros which have the same first 3 letters
             then all of them will be executed by Zemax.
        """
        status = -1
        if self.macroPath:
            zplMpath = self.macroPath
        else:
            zplMpath = _os.path.join(self.zGetPath()[0], 'Macros')
        macroList = [f for f in _os.listdir(zplMpath)
                     if f.endswith(('.zpl','.ZPL')) and f.startswith(zplMacroCode)]
        if macroList:
            zplCode = macroList[0][:3]
            status = self.zOpenWindow(zplCode, True, timeout)
        return status

    def zExportCAD(self, fileName, fileType=1, numSpline=32, firstSurf=1,
                   lastSurf=-1, raysLayer=1, lensLayer=0, exportDummy=0,
                   useSolids=1, rayPattern=0, numRays=0, wave=0, field=0,
                   delVignett=1, dummyThick=1.00, split=0, scatter=0,
                   usePol=0, config=0):
        """Export lens data in IGES/STEP/SAT format for import into CAD programs

        Parameters
        ----------
        fileName : string
            filename including extension (including full path is recommended)
        fileType : integer (0, 1, 2 or 3)
            0 = IGES; 1 = STEP (default); 2 = SAT; 3 = STL
        numSpline : integer
            number of spline segments to use (default = 32)
        firstSurf : integer
            the first surface to export; the first object to export (in NSC mode)
        lastSurf : integer
            the last surface to export; the last object to export (in NSC mode)
            (default = -1, i.e. image surface)
        raysLayer : integer
            layer to place ray data on (default = 1)
        lensLayer : integer
            layer to place lens data on (default = 0)
        exportDummy : integer (0 or 1)
            export dummy surface? 1 = export; 0 (default) = not export
        useSolids : integer (0 or 1)
            export surfaces as solids? 1 (default) = surfaces as solids;
        rayPattern : integer (0 <= rayPattern <= 7)
            0 (default) = XY fan; 1 = X fan; 2 = Y fan; 3 = ring; 4 = list;
            5 = none; 6 = grid; 7 = solid beams.
        numRays : integer
            the number of rays to render (default = 1)
        wave : integer
            wavelength number; 0 (default) indicates all
        field : integer
            the field number; 0 (default) indicates all
        delVignett : integer (0 or 1)
            delete vignetted rays? 1 (default) = delete vig. rays
        dummyThick : float
            dummy surface thickness in lens units; (default = 1.00)
        split : integer (0 or 1)
            split rays from NSC sources? 1 = split sources; 0 (default) = no
        scatter : integer (0 or 1)
            scatter rays from NSC sources? 1 = Scatter; 0 (deafult) = no
        usePol : integer (0 or 1)
            use polarization when tracing NSC rays? 1 = use polarization;
            0 (default) no. Note that polarization is automatically selected
            if ``split`` is ``1``.
        config : integer (0 <= config <= n+3)
            n is the total number of configurations;
            0 (default) = current config;
            1 - n for a specific configuration;
            n+1 to export "All By File";
            n+2 to export "All by Layer";
            n+3 for "All at Once".

        Returns
        -------
        status : string
            the string "Exporting filename" or "BUSY!" (see notes below)

        Notes
        -----
        1. Exporting lens data data may take longer than the timeout interval of
           the DDE communication. Zemax spwans an independent thread to process
           this request. Once the thread is launched, Zemax returns
           "Exporting filename". However, the export may take much longer.
           To verify the completion of export and the readiness of the file,
           use ``zExportCheck()``, which returns ``1`` as long as the export is
           in process, and ``0`` once completed. Generally, ``zExportCheck()``
           should be placed in a loop, which executes until a ``0`` is returned.

           A typical loop-test may look like as follows: ::

            # check for completion of CAD export process
            still_working = True
            while(still_working):
                # Delay for 200 milliseconds
                time.sleep(.2)
                status = ln.zExportCheck()
                if status:  # still running
                    pass
                else:       # Done exporting
                    still_working = False

        2. Zemax cannot export some NSC objects (e.g. slide). The unexportable
           objects are ignored.

        References
        ----------
        For a detailed exposition on the configuration settings,
        see "Export IGES/SAT.STEP Solid" in the Zemax manual.
        """
        # Determine last surface/object depending upon zemax mode
        if lastSurf == -1:
            zmxMode = self._zGetMode()
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
        reply = self._sendDDEcommand(cmd)
        return str(reply)

    def zExportCheck(self):
        """Indicate the status of the last executed ``zExportCAD()`` command

        Parameters
        ----------
        None

        Returns
        -------
        status : integer (0 or 1)
            0 = last CAD export completed; 1 = last CAD export in progress
        """
        return int(self._sendDDEcommand('ExportCheck'))

    def zFindLabel(self, label):
        """Returns the surface that has the integer label associated with the it.

        Parameters
        ----------
        label : integer
            label associated with a surface

        Returns
        -------
        surfaceNumber : integer
            surface-number of surface associated with the given ``label``;
            -1 if no surface with the specified label is found.

        See Also
        --------
        zSetLabel(), zGetLabel()
        """
        reply = self._sendDDEcommand("FindLabel,{:d}".format(label))
        return int(float(reply))

    def zGetAddress(self, addressLineNumber):
        """Extract the address in specified line number

        Parameters
        ----------
        addressLineNumber : integer
            line number of address to return

        Returns
        -------
        addressLine : string
            address line
        """
        reply = self._sendDDEcommand("GetAddress,{:d}"
                                    .format(addressLineNumber))
        return str(reply)

    def zGetAperture(self, surfNum):
        """Get the surface aperture data for a given surface

        Parameters
        ----------
        surfNum : integer
            the surface-number of a surface

        Returns
        -------
        aType : integer
            integer codes of aperture types which are:

                * 0 = no aperture (na);
                * 1 = circular aperture (ca);
                * 2 = circular obscuration (co);
                * 3 = spider (s);
                * 4 = rectangular aperture (ra);
                * 5 = rectangular obscuration (ro);
                * 6 = elliptical aperture (ea);
                * 7 = elliptical obscuration (eo);
                * 8 = user defined aperture (uda);
                * 9 = user defined obscuration (udo);
                * 10 = floating aperture (fa);

        aMin : float
            min radius(ca); min radius(co); width of arm(s); X-half width(ra);
            X-half width(ro); X-half width(ea); X-half width(eo)
        aMax : float
            max radius(ca); max radius(co); number of arm(s); X-half width(ra);
            X-half width(ro); X-half width(ea); X-half width(eo)
        xDecenter : float
            amount of decenter in x from current optical axis (in lens units)
        yDecenter : float
            amount of decenter in y from current optical axis (in lens units)
        apertureFile : string
            a text file with .UDA extention.

        References
        ----------
        The following sections from the Zemax manual should be referred for
        details:

        1. "Aperture type and other aperture controls" for details on aperture
        2. "User defined apertures and obscurations" for more on UDA extension

        See Also
        --------
        zGetSystemAper() :
            For system aperture instead of the aperture of surface.
        zSetAperture()
        """
        reply = self._sendDDEcommand("GetAperture,"+str(surfNum))
        rs = reply.split(',')
        apertureInfo = [int(rs[i]) if i==5 else float(rs[i])
                                             for i in range(len(rs[:-1]))]
        apertureInfo.append(rs[-1].rstrip()) # append the test file (string)
        return tuple(apertureInfo)

    def zGetApodization(self, px, py):
        """Computes the intensity apodization of a ray from the apodization
        type and value.

        Parameters
        ----------
        px, py : float
            normalized pupil coordinates

        Returns
        -------
        intApod : float
            intensity apodization
        """
        reply = self._sendDDEcommand("GetApodization,{:1.20g},{:1.20g}"
                                    .format(px,py))
        return float(reply)

    def zGetAspect(self, filename=None):
        """Returns the graphic display aspect-ratio and the width (or height)
        of the printed page in current lens units.

        Parameters
        ----------
        filename : string
            name of the temporary file associated with the window being
            created or updated. If the temporary file is left off, then the
            default aspect-ratio and width (or height) is returned.

        Returns
        -------
        aspect : float
            aspect ratio (height/width)
        side : float
            width if ``aspect <= 1``; height if ``aspect > 1`` (in lens units)
        """
        asd = _co.namedtuple('aspectData', ['aspect', 'side'])
        cmd = (filename and "GetAspect,{}".format(filename)) or "GetAspect"
        reply = self._sendDDEcommand(cmd).split(",")
        aspectSide = asd._make([float(elem) for elem in reply])
        return aspectSide

    def zGetBuffer(self, n, tempFileName):
        """Retrieve DDE client specific data from a window being updated

        Parameters
        ----------
        n : integer (0 <= n <= 15)
            the buffer number
        tempFileName : string
            name of the temporary file associated with the window being
            updated. The tempFileName is passed to the client when Zemax
            calls the client.

        Returns
        -------
        bufferData : string
            buffer data

        Notes
        -----
        Each window may have its own buffer data, and Zemax uses the filename
        to identify the window for which the buffer data is requested.

        References
        ----------
        See section "How ZEMAX calls the client" in Zemax manual.

        See Also
        --------
        zSetBuffer()
        """
        cmd = "GetBuffer,{:d},{}".format(n,tempFileName)
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())
        # !!!FIX what is the proper return for this command?

    def zGetComment(self, surfaceNumber):
        """Returns the surface comment, if any, associated with the surface

        Parameters
        ----------
        surfaceNumber: integer
            the surface number

        Returns
        -------
        comment : string
            the comment, if any, associated with the surface
        """
        reply = self._sendDDEcommand("GetComment,{:d}".format(surfaceNumber))
        return str(reply.rstrip())

    def zGetConfig(self):
        """Returns tuple containing current configuration number, number of
        configurations, and number of multiple configuration operands.

        Parameters
        ----------
        none

        Returns
        -------
        currentConfig : integer
            current configuration (column) number in MCE
        numberOfConfigs : integer
            number of configurations (number of columns)
        numberOfMutiConfigOper : integer
            number of multi config operands (number of rows)

        Notes
        -----
        The function returns ``(1,1,1)`` even if the multi-configuration editor
        is empty. This is because, the current lens in the LDE is, by default,
        set to the current configuration. The initial number of configurations
        is therefore ``1``, and the number of operators in the
        multi-configuration editor is also ``1`` (usually, ``MOFF``).

        See Also
        --------
        zInsertConfig() :
            Use ``zInsertConfig()`` to insert new configuration in the
            multi-configuration editor.
        zSetConfig()
        """
        reply = self._sendDDEcommand('GetConfig')
        rs = reply.split(',')
        # !!! FIX: Should this function return "0" when the MCE is empty, just
        # like what is done for the zGetNSCData() function?
        return tuple([int(elem) for elem in rs])

    def zGetDate(self):
        """Get current date from the Zemax DDE server.

        Parameters
        ----------
        None

        Returns
        -------
        date : string
            date
        """
        return self._sendDDEcommand('GetDate')

    def zGetExtra(self, surfaceNumber, columnNumber):
        """Returns extra surface data from the Extra Data Editor

        Parameters
        ----------
        surfaceNumber : integer
            surface number
        columnNumber : integer
            column number

        Returns
        -------
        value : float
            numeric data value

        See Also
        --------
        zSetExtra()
        """
        cmd="GetExtra,{sn:d},{cn:d}".format(sn=surfaceNumber,cn=columnNumber)
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zGetField(self, n):
        """Returns field data from Zemax DDE server

        Parameters
        ----------
        n : integer
            field number

        Returns
        -------
        fieldData : tuple
            exact elements of the tuple depeds on the value of ``n``

            if ``n = 0``
                fieldData is a tuple containing the following:

                  - type : integer (0 = angles in degrees, \
                                    1 = object height, \
                                    2 = paraxial image height, \
                                    3 = real image height)
                  - numFields : number of fields currently defined
                  - max_x_field : value used to normalize x field coordinate
                  - max_y_field : value used to normalize y field coordinate
                  - normalization_method : field normalization method
                    (0 = radial, 1 = rectangular)

            if ``0 < n <= number_of_fields``
                fieldData is a tuple containing the following:

                  - x : the field x value
                  - y : the field y value
                  - wt : field weight
                  - vdx : decenter x vignetting factor
                  - vdy : decenter y vignetting factor
                  - vcx : compression x vignetting factor
                  - vcy : compression y vignetting factor
                  - van : angle vignetting factor

        Notes
        -----
        The returned tuple's content and structure is exactly same as that
        returned by ``zSetField()``

        See Also
        --------
        zSetField()
        """
        if n: # n > 0
            fd = _co.namedtuple('fieldData', ['X', 'Y', 'wt',
                                              'vdx', 'vdy',
                                              'vcx', 'vcy', 'van'])
        else: # n = 0
            fd = _co.namedtuple('fieldData', ['type', 'numFields',
                                              'Xmax', 'Ymax', 'normMethod'])
        reply = self._sendDDEcommand('GetField,'+ str(n))
        rs = reply.split(',')
        if n: # n > 0
            fieldData = fd._make([float(elem) for elem in rs])
        else: # n = 0
            fieldData = fd._make([int(elem) if (i==0 or i==1)
                                 else float(elem) for i,elem in enumerate(rs)])
        return fieldData

    def zGetFieldTuple(self):
        """Get all field data in a single n-tuple.

        Parameters
        ----------
        None

        Returns
        -------
        fieldDataTuple: n-tuple (0 < n <= 12)
            the tuple elements represent field loactions with each element
            containing all 8 field parameters.

        Examples
        --------
        >>> # example shows the namedtuple returned by ``zGetFieldTuple``
        >>> ln.zGetFieldTuple()
        (fieldData(X=0.0, Y=0.0, wt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0),
         fieldData(X=0.0, Y=14.0, wt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0),
         fieldData(X=0.0, Y=20.0, wt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0))

        See Also
        --------
        zGetField(), zSetField(), zSetFieldTuple()
        """
        fieldCount = self.zGetField(0)[1]
        fd = _co.namedtuple('fieldData', ['X', 'Y', 'wt',
                                          'vdx', 'vdy',
                                          'vcx', 'vcy', 'van'])
        fieldData = [ ]
        for i in range(fieldCount):
            reply = self._sendDDEcommand('GetField,'+str(i+1))
            rs = reply.split(',')
            data = fd._make([float(elem) for elem in rs])
            fieldData.append(data)
        return tuple(fieldData)

    def zGetFile(self):
        """Returns the full name of the zmx lens file in Zemax server

        Parameters
        ----------
        None

        Returns
        -------
        file_name : string
            filename of the zmx file in the Zemax DDE server.

        Notes
        -----
        Extreme caution should be used if the file is to be tampered with;
        at any time Zemax may read or write from/to this file.
        """
        reply = self._sendDDEcommand('GetFile')
        return reply.rstrip()

    def zGetFirst(self):
        """Returns the first order lens data

        Parameters
        ----------
        None

        Returns
        -------
        focal : float
            Effective Focal Length (EFL) in lens units,
        pwfn : float
            paraxial working F/#,
        rwfn : float
            real working F/#,
        pima : float
            paraxial image height, and
        pmag : float
            paraxial magnification.

        See Also
        --------
        zGetSystem() :
            Use ``zGetSystem()`` to get general system data,
        zGetSystemProperty()
        """
        fd = _co.namedtuple('firstOrderData',
                            ['EFL', 'paraWorkFNum', 'realWorkFNum',
                             'paraImgHeight', 'paraMag'])
        reply = self._sendDDEcommand('GetFirst')
        rs = reply.split(',')
        return fd._make([float(elem) for elem in rs])

    def zGetGlass(self, surfaceNumber):
        """Returns glass data of a surface.

        Parameters
        ----------
        surfaceNumber : integer
            surface number

        Returns
        -------
        glassInfo : 4-tuple or None
            glassInfo contains (name, nd, vd, dpgf) if there is a valid glass
            associated with the surface, else ``None``

        Notes
        -----
        If the specified surface is not valid, is not made of glass, or is
        gradient index, the returned string is empty. This data may be
        meaningless for glasses defined only outside of the FdC band.
        """
        gd = _co.namedtuple('glassData', ['name', 'nd', 'vd', 'dpgf'])
        reply = self._sendDDEcommand("GetGlass,{:d}".format(surfaceNumber))
        rs = reply.split(',')
        if len(rs) > 1:
            glassInfo = gd._make([str(rs[i]) if i == 0 else float(rs[i])
                                                      for i in range(len(rs))])
        else:
            glassInfo = None
        return glassInfo

    def zGetGlobalMatrix(self, surfaceNumber):
        """Returns the the matrix required to convert any local coordinates
        (such as from a ray trace) into global coordinates.

        Parameters
        ----------
        surfaceNumber : integer
            surface number

        Returns
        -------
        globalMatrix : 9-tuple
            the elements of the global matrix:
        |        (R11, R12, R13,
        |         R21, R22, R23,
        |         R31, R32, R33,
        |         Xo,  Yo , Zo)

        the function returns -1, if bad command.

        References
        ----------
        For details on the global coordinate matrix, see "Global Coordinate
        Reference Surface" in the Zemax manual.
        """
        gmd = _co.namedtuple('globalMatrix', ['R11', 'R12', 'R13',
                                              'R21', 'R22', 'R23',
                                              'R31', 'R32', 'R33',
                                              'Xo' ,  'Yo', 'Zo'])
        cmd = "GetGlobalMatrix,{:d}".format(surfaceNumber)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        globalMatrix = gmd._make([float(elem) for elem in rs.split(',')])
        return globalMatrix

    def zGetIndex(self, surfaceNumber):
        """Returns the index of refraction data for the specified surface

        Parameters
        ----------
        surfaceNumber : integer
            surface number

        Returns
        -------
        indexData : tuple of real values
            the ``indexData`` is a tuple of index of refraction values
            defined for each wavelength in the format (n1, n2, n3, ...).
            If the specified surface is not valid, or is gradient index,
            the returned string is empty.
        """
        reply = self._sendDDEcommand("GetIndex,{:d}".format(surfaceNumber))
        rs = reply.split(",")
        indexData = [float(rs[i]) for i in range(len(rs))]
        return tuple(indexData)

    def zGetLabel(self, surfaceNumber):
        """Returns the integer label associated with the specified surface.

        Parameters
        ----------
        surfaceNumber : integer
            the surface number

        Returns
        -------
        label : integer
            the integer label

        Notes
        -----
        Labels are retained by Zemax as surfaces are inserted or deleted
        around the target surface.

        See Also
        --------
        zSetLabel(), zFindLabel()
        """
        reply = self._sendDDEcommand("GetLabel,{:d}".format(surfaceNumber))
        return int(float(reply.rstrip()))

    def zGetMetaFile(self, metaFileName, analysisType, settingsFile=None,
                     flag=0):
        """Creates a windows Metafile of any Zemax graphical analysis window

        Usage: ``zMetaFile(metaFilename, analysisType [, settingsFile, flag])``

        Parameters
        ----------
        metaFileName : string
            absolute path name with extension
        analysisType : string
            3-letter case-sensitive button code for the analysis. If no label
            is provided or recognized, a 3D Layout plot is generated.
        settingsFile : string
            settings file used/ saved by Zemax to compute the metafile graphic
            depending upon the value of the flag parameter.
        flag : integer
            0 = default settings used for the graphic;
            1 = use settings in settings file if valid, else default settings;
            2 = use settings in settings file if valid, and the settings box
            will be displayed for further setting changes.

        Returns
        -------
        status : integer
            0  = Success;
            -1 = metafile could not be saved;
            -998 = command timed out

        Notes
        -----
        No matter what the flag value is, if a valid file-name is provided
        for the ``settingsFile``, the settings used will be written to
        the settings file, overwriting any data in the file.

        Examples
        --------
        >>> ln.zGetMetaFile("C:\\Projects\\myGraphicfile.EMF",'Lay')
        0

        See Also
        --------
        zGetTextFile(), zOpenWindow()
        """
        if settingsFile:
            settingsFile = settingsFile
        else:
            settingsFile = ''
        retVal = -1
        # Check if Valid analysis type
        if zb.isZButtonCode(analysisType):
            # Check if the file path is valid and has extension
            if _os.path.isabs(metaFileName) and _os.path.splitext(metaFileName)[1]!='':
                cmd = 'GetMetaFile,"{tF}",{aT},"{sF}",{fl:d}'.format(tF=metaFileName,
                                    aT=analysisType,sF=settingsFile,fl=flag)
                reply = self._sendDDEcommand(cmd)
                if 'OK' in reply.split():
                    retVal = 0
        else:
            print("Invalid analysis code '{}' passed to zGetMetaFile."
                  .format(analysisType))
        return retVal

    def _zGetMode(self):
        """Returns the mode (Sequential, Non-sequential or Mixed) of the current
        lens in the DDE server

        Parameters
        ----------
        None

        Returns
        -------
        zmxModeInformation : 2-tuple (mode, nscSurfNums)
            mode (0 = Sequential; 1 = Non-sequential; 2 = Mixed mode)
            nscSurfNums = (tuple of integers) the surfaces (in mixed mode)
            that are non-sequential. In Non-sequential mode and in purely
            sequential mode, this tuple is empty (of length 0).

        Notes
        -----
        This only works when a zmx file is loaded into the server. Currently this
        function is meant to be used for internal purpose only.

        For the purpose of this function, "Sequential" implies that there are no
        non-sequential surfaces in the LDE.
        """
        nscSurfNums = []
        nscData = self.zGetNSCData(1, 0)
        if nscData > 0: # Non-sequential mode
            mode = 1
        else:          # Not Non-sequential mode
            numSurf = self.zGetSystem()[0]
            for i in range(1,numSurf+1):
                surfType = self.zGetSurfaceData(i, 0)
                if surfType == 'NONSEQCO':
                    nscSurfNums.append(i)
            if len(nscSurfNums) > 0:
                mode = 2  # mixed mode
            else:
                mode = 0  # sequential
        return (mode,tuple(nscSurfNums))

    def zGetMulticon(self, config, row):
        """Returns data from the multi-configuration editor

        Parameters
        ----------
        config : integer
            configuration number (column)
        row : integer
            operand

        Returns
        -------
        multiConData : tuple
            the exact elements of ``multiConData`` depends on the value of
            ``config``

            If ``config > 0``
                then the elements of ``multiConData`` are:
                (value, num_config, num_row, status, pickuprow, pickupconfig,
                scale, offset)

                The ``status`` is 0 for fixed, 1 for variable, 2 for pickup,
                & 3 for thermal pickup. If ``status`` is 2 or 3, the pickuprow &
                pickupconfig values indicate the source data for the pickup solve.

            If ``config = 0``
                then the elements of ``multiConData`` are:
                (operand_type, number1, number2, number3)

        See Also
        --------
        zSetMulticon()
        """
        cmd = "GetMulticon,{config:d},{row:d}".format(config=config,row=row)
        reply = self._sendDDEcommand(cmd)
        if config: # if config > 0
            rs = reply.split(",")
            if '' in rs: # if the MCE is "empty"
                rs[rs.index('')] = '0'
            multiConData = [float(rs[i]) if (i==0 or i==6 or i==7) else int(rs[i])
                                                         for i in range(len(rs))]
        else: # if config == 0
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0, rs[0])
        return tuple(multiConData)

    def zGetName(self):
        """Returns the name of the lens

        Parameters
        ----------
        None

        Returns
        -------
        lensName : string
            name of the current lens (as in the General data dialog box)
        """
        reply = self._sendDDEcommand('GetName')
        return str(reply.rstrip())

    def zGetNSCData(self, surfaceNumber, code):
        """Returns the data for NSC groups

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group; Use 1 if for pure NSC mode
        code : integer (0)
            currently only ``code = 0`` is supported, in which case the
            returned data is the number of objects in the NSC group

        Returns
        -------
        nscData :
            the number of objects in the NSC group if the command is valid;
            -1 if it was a bad commnad (generally if the ``surface`` is not
            a non-sequential surface)

        Notes
        -----
        This function returns 1 if the only object in the NSC editor is a
        "Null Object".
        """
        cmd = "GetNSCData,{:d},{:d}".format(surfaceNumber,code)
        reply = self._sendDDEcommand(cmd)
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

    def zGetNSCMatrix(self, surfaceNumber, objectNumber):
        """Returns a tuple containing the rotation and position matrices
        relative to the NSC surface origin.

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group; Use 1  for pure NSC mode
        objectNumber : integer
            the NSC ojbect number

        Returns
        -------
        nscMatrix : 9-tuple
            the elements of the global matrix:
        |        (R11, R12, R13,
        |         R21, R22, R23,
        |         R31, R32, R33,
        |         Xo,  Yo , Zo)

        the function returns -1, if bad command.
        """
        nscmat = _co.namedtuple('NSCMatrix', ['R11', 'R12', 'R13',
                                              'R21', 'R22', 'R23',
                                              'R31', 'R32', 'R33',
                                              'Xo' ,  'Yo', 'Zo'])
        cmd = "GetNSCMatrix,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscMatrix = -1
        else:
            nscMatrix = nscmat._make([float(elem) for elem in rs.split(',')])
        return nscMatrix

    def zGetNSCObjectData(self, surfaceNumber, objectNumber, code):
        """Returns the various data for NSC objects.

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        code : integer
            for the specific code see the nsc-object-data-codes_ table (below)

        Returns
        -------
        nscObjectData : string/integer/float
            the nature of the returned data, which depends on the ``code``,
            is enumerated in the nsc-object-data-codes_  table (below).
            If the command fails, it returns ``-1``.

        Notes
        -----

        .. _nsc-object-data-codes:

        ::

            Table: Codes for NSC object data getter and setter methods

            --------------------------------------------------------------------
            code - Datum set/returned by zSetNSCObjectData()/zGetNSCObjectData()
            --------------------------------------------------------------------
              0  - Object type name (string).
              1  - Comment and/or the file name if the object is defined by a
                   file (string).
              2  - Color (integer).
              5  - Reference object number (integer).
              6  - Inside of object number (integer).

            These codes set/get values to/from the "Type tab" of the Object
            Properties dialog:
              3  - 1 if object uses a user defined aperture file, 0 otherwise
              4  - User defined aperture file name, if any (string).
             29  - "Use Pixel Interpolation" checkbox (1 = checked, 0 = unchecked).

            These codes set/get values to/from the "Sources tab" of the Object
            Properties dialog:
            101  - Source object random polarization (1 = checked, 0 = unchecked)
            102  - Source object reverse rays option (1 = checked, 0 for unchecked)
            103  - Source object Jones X value.
            104  - Source object Jones Y value.
            105  - Source object Phase X value.
            106  - Source object Phase Y value.
            107  - Source object initial phase in degrees value.
            108  - Source object coherence length value.
            109  - Source object pre-propagation value.
            110  - Source object sampling method (0 = random, 1 = Sobol sampling)
            111  - Source object bulk scatter method (0 = many, 1 = once, 2 = never)

            These codes set/set values to/from the "Bulk Scatter tab" of the Object
            Properties dialog:
            202  - Mean Path value.
            203  - Angle value.
            211-226 - DLL parameter 1-16, respectively.

            end-of-table

        See Also
        --------
        zSetNSCObjectData()
        """
        str_codes = (0,1,4)
        int_codes = (2,3,5,6,29,101,102,110,111)
        cmd = ("GetNSCObjectData,{:d},{:d},{:d}"
              .format(surfaceNumber,objectNumber,code))
        reply = self._sendDDEcommand(cmd)
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

    def zGetNSCObjectFaceData(self, surfNumber, objNumber, faceNumber, code):
        """Returns the various data for NSC object faces.

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        faceNumber : integer
            face number
        code : integer
            code (see below)

        Returns
        -------
        nscObjFaceData  : data for NSC object faces (see the table for the
                          particular type of data) if successful, else -1

        Notes
        -----

        .. _nsc-object-face-data-codes:

        ::

            Table: Codes for NSC object face data getter and setter methods

            --------------------------------------------------------------------
            code  -  Datum set/ret by zGetNSCObjectFaceData/zGetNSCObjectFaceData
            --------------------------------------------------------------------
             10   -  Coating name (string).
             20   -  Scatter code (0 = None, 1 = Lambertian, 2 = Gaussian,
                     3 = ABg, and 4 = user defined)
             21   -  Scatter fraction (float).
             22   -  Number of rays to scatter (integer).
             23   -  Gaussian scatter sigma (float).
             24   -  Face is setting(0 = object default, 1 = reflective,
                     2 = absorbing)
             30   -  ABg scatter profile name for reflection (string).
             31   -  ABg scatter profile name for transmission (string).
             40   -  User Defined Scatter DLL name (string).
             41-46 - User Defined Scatter Parameter 1 - 6 (double).
             60   -  User Defined Scatter data file name (string).

             end-of-table

        See Also
        --------
        zSetNSCObjectFaceData()
        """
        str_codes = (10,30,31,40,60)
        int_codes = (20,22,24)
        cmd = ("GetNSCObjectFaceData,{:d},{:d},{:d},{:d}"
              .format(surfNumber,objNumber,faceNumber,code))
        reply = self._sendDDEcommand(cmd)
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

    def zGetNSCParameter(self, surfNumber, objNumber, parameterNumber):
        """Returns NSC object's parameter data

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        parameterNumber : integer
            parameter number

        Returns
        -------
        nscParaVal : float
            parameter value

        See Also
        --------
        zSetNSCParameter()
        """
        cmd = ("GetNSCParameter,{:d},{:d},{:d}"
              .format(surfNumber,objNumber,parameterNumber))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zGetNSCPosition(self, surfNumber, objectNumber):
        """Returns position data for NSC object

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number

        Returns
        -------
        nscPos : 7-tuple (x, y, z, tilt-x, tilt-y, tilt-z, material)

        Examples
        --------
        >>> ln.zGetNSCPosition(surfNumber=1, objectNumber=4)
        NSCPosition(x=0.0, y=0.0, z=10.0, tiltX=0.0, tiltY=0.0, tiltZ=0.0, material='N-BK7')

        See Also
        --------
        zSetNSCPosition()
        """
        nscpd = _co.namedtuple('NSCPosition', ['x', 'y', 'z',
                                               'tiltX', 'tiltY', 'tiltZ',
                                               'material'])
        cmd = ("GetNSCPosition,{:d},{:d}".format(surfNumber,objectNumber))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPos = -1
        else:
            nscPos = nscpd._make([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                      for i in range(len(rs))])
        return nscPos

    def zGetNSCProperty(self, surfaceNumber, objectNumber, faceNumber, code):
        """Returns a numeric or string value from the property pages of objects
        defined in NSC editor. It mimics the ZPL function NPRO.

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        faceNumber : integer
            face number
        code : integer
            for the specific code see the nsc-property-codes_ table (below)

        Returns
        -------
        nscPropData : string/float/integer
            the nature of the returned data, which depends on the ``code``,
            is enumerated in the nsc-property-codes_  table (below).
            If the command fails, it returns ``-1``.

        Notes
        -----

        .. _nsc-property-codes:

        ::

            Table: Codes for NSC property getter and setter methods

            --------------------------------------------------------------------
            code - Datum set/returned by zSetNSCProperty()/zGetNSCProperty()
            --------------------------------------------------------------------
            The following codes sets/get values to/from the NSC Editor.
              1 - Object comment (string).
              2 - Reference object number (integer).
              3 - "Inside of" object number (integer).
              4 - Object material (string).

            The following codes set/get values to/from the "Type tab" of the
            Object Properties dialog.
              0 - Object type. For e.g., "NSC_SLEN" for the standard lens (string).
             13 - User Defined Aperture (1 = checked, 0 = unchecked)
             14 - User Defined Aperture file name (string).
             15 - "Use Global XYZ Rotation Order" checkbox; (1 = checked,
                  0 = unchecked)
             16 - "Rays Ignore This Object" checkbox; (1 = checked, 0 = un-checked)
             17 - "Object Is A Detector" checkbox; (1 = checked, 0 = un-checked)
             18 - "Consider Objects" list. The argument is a string listing the
                  object numbers delimited by spaces, e.g., "2 5 14" (string).
             19 - "Ignore Objects" list. The argument is a string listing the
                  object numbers delimited by spaces, e.g., "1 3 7" (string).
             20 - "Use Pixel Interpolation" checkbox, (1 = checked, 0 = un-
                  checked).

            The following codes set/get values to/from the "Coat/Scatter tab" of
            the Object Properties dialog.
              5 - Coating name for the specified face (string).
              6 - Profile name for the specified face (string).
              7 - Scatter mode for the specified face, (0 = none, 1 = Lambertian,
                  2 = Gaussian, 3 = ABg, 4 = User Defined.)
              8 - Scatter fraction for the specified face (float).
              9 - Number of scatter rays for the specified face (integer).
             10 - Gaussian sigma for the specified face (float).
             11 - Reflect ABg data name for the specified face (string).
             12 - Transmit ABg data name for the specified face (string).
             27 - Name of the user defined scattering DLL (string).
             28 - Name of the user defined scattering data file (string).
            21-26 - Parameter values on the user defined scattering DLL (float).
             29 - "Face Is" property for the specified face (0 = "Object Default",
                  1 = "Reflective", 2 = "Absorbing")

            The following codes set/get values to/from the "Bulk Scattering tab" of
            the Object Properties dialog.
             81 - "Model" value on the bulk scattering tab (0 = "No Bulk
                  Scattering", 1 = "Angle Scattering", 2 = "DLL Defined Scattering")
             82 - Mean free path to use for bulk scattering.
             83 - Angle to use for bulk scattering.
             84 - Name of the DLL to use for bulk scattering.
             85 - Parameter value to pass to the DLL, where the face value
                  is used to specify which parameter is being defined. The first
                  parameter is 1, the second is 2, etc. (float)
             86 - Wavelength shift string (string).

            The following codes set/get values from the Diffraction tab of the
            Object Properties dialog.
             91 - "Split" value on the diffraction tab (0 = "Don't Split By Order",
                  1 = "Split By Table Below", 2 = "Split By DLL Function")
             92 - Name of the DLL to use for diffraction splitting (string).
             93 - Start Order value (float).
             94 - Stop Order value (float).
             95 - Parameter values on the diffraction tab. These parameters are
                  passed to the diffraction splitting DLL as well as the order
                  efficiency values used by "split by table below" option. The
                  face value is used to specify which parameter is being defined.
                  The first parameter is 1, the second is 2, etc. (float).

            The following codes set/get values to/from the "Sources tab" of the
            Object Properties dialog.
            101 - Source object random polarization (1=checked, 0=unchecked).
            102 - Source object reverse rays option (1=checked, 0=unchecked).
            103 - Source object Jones X value.
            104 - Source object Jones Y value.
            105 - Source object Phase X value.
            106 - Source object Phase Y value.
            107 - Source object initial phase in degrees value.
            108 - Source object coherence length value.
            109 - Source object pre-propagation value.
            110 - Source object sampling method; (0=random, 1=Sobol sampling)
            111 - Source object bulk scatter method; (0=many,1=once, 2=never)
            112 - Array mode; (0 = none, 1 = rectangular, 2 = circular,
                  3 = hexapolar, 4 = hexagonal)
            113 - Source color mode. For a complete list of the available
                  modes, see "Defining the color and spectral content of sources"
                  in the Zemax manual. The source color modes are numbered starting
                  with 0 for the System Wavelengths, and then from 1 through the last
                  model listed in the dialog box control (integer).
            114-116 - Number of spectrum steps, start wavelength, and end
                      wavelength, respectively (float).
            117 - Name of the spectrum file (string).
            161-162 - Array mode integer arguments 1 and 2.
            165-166 - Array mode double precision arguments 1 and 2.
            181-183 - Source color mode arguments, for example, the XYZ
                      values of the Tristimulus (float).

            The following codes set/get values to/from the "Grin tab" of the
            Object Properties dialog.
            121 - "Use DLL Defined Grin Media" checkbox (1 = checked, 0 =
                  unchecked).
            122 - Maximum step size value (float).
            123 - DLL name (string).
            124 - Grin DLL parameters. These are the parameters passed to the DLL.
                  The face value is used to specify which parameter is being
                  defined. The first parameter is 1, the second is 2, etc (float)

            The following codes set/get values to/from the "Draw tab" of the
            Object Properties dialog.
            141 - Do not draw object checkbox (1 = checked, 0 = unchecked)
            142 - Object opacity (0 = 100%, 1 = 90%, 2 = 80%, etc.)

            The following codes set/get values to/from the "Scatter To tab" of
            the Object Properties dialog.
            151 - Scatter to method (0 = scatter to list, 1 = importance
                  sampling)
            152 - Importance Sampling target data. The argument is a string
                  listing the ray number, the object number, the size, & the
                  limit value, separated by spaces. For e.g., to set the
                  Importance Sampling data for ray 3, object 6, size 3.5, and
                  limit 0.6, the string argument is "3 6 3.5 0.6".
            153 - "Scatter To List" values. The argument is a string listing the
                  object numbers to scatter to delimited by spaces, such as
                  "4 6 19" (string).

            The following codes set/get values to/from the "Birefringence tab"
            of the Object Properties dialog.
            171 - Birefringent Media checkbox (0 = unchecked, 1 = checked)
            172 - Birefringent Media Mode (0 = Trace ordinary and extraordinary
                  rays, 1 = Trace only ordinary rays, 2 = Trace only
                  extraordinary rays, and 3 = Waveplate mode)
            173 - Birefringent Media Reflections status (0 = Trace reflected and
                  refracted rays, 1 = Trace only refracted rays, and 2 = Trace
                  only reflected rays)
            174-176 - Ax, Ay, and Az values (float).
            177 - Axis Length (float).
            200 - Index of refraction of an object (float).
            201-203 - nd (201), vd (202), and dpgf (203) parameters of an
                      object using a model glass.

            end-of-table

        See Also
        --------
        zSetNSCProperty()
        """
        cmd = ("GetNSCProperty,{:d},{:d},{:d},{:d}"
                .format(surfaceNumber, objectNumber, code, faceNumber))
        reply = self._sendDDEcommand(cmd)
        nscPropData = _process_get_set_NSCProperty(code, reply)
        return nscPropData

    def zGetNSCSettings(self):
        """Returns the maximum number of intersections, segments, nesting level,
        minimum absolute intensity, minimum relative intensity, glue distance,
        miss ray distance, and ignore errors flag used for NSC ray tracing.

        Parameters
        ----------
        None

        Returns
        -------
        maxInt : integer
            maximum number of intersections
        maxSeg : integer
            maximum number of segments
        maxNest : integer
            maximum nesting level
        minAbsI : float
            minimum absolute intensity
        minRelI : float
            minimum relative intensity
        glueDist : float
            glue distance
        missRayLen : float
            miss ray distance
        ignoreErr : integer
            1 if true, 0 if false

        See Also
        --------
        zSetNSCSettings()
        """
        reply = str(self._sendDDEcommand('GetNSCSettings'))
        rs = reply.rsplit(",")
        nscSettingsData = [float(rs[i]) if i in (3,4,5,6) else int(float(rs[i]))
                                                        for i in range(len(rs))]
        return tuple(nscSettingsData)

    def zGetNSCSolve(self, surfaceNumber, objectNumber, parameter):
        """Returns the current solve status and settings for NSC position and
        parameter data

        Parameters
        ----------
        surfaceNumber : integer
            surface number of NSC group; use 1 if the program mode is pure NSC
        objectNumber : integer
            object number
        parameter : integer
            the parameter are as follows:

                * -1 = extract data for x data
                * -2 = extract data for y data
                * -3 = extract data for z data
                * -4 = extract data for tilt x data
                * -5 = extract data for tilt y data
                * -6 = extract data for tilt z data
                * n > 0  = extract data for the nth parameter

        Returns
        -------
        nscSolveData : 5-tuple
            nscSolveData tuple contains
            (status, pickupObject, pickupColumn, scaleFactor, offset)

            The status value is 0 for fixed, 1 for variable, and 2 for a pickup
            solve.

            Only when the staus is a pickup solve is the other data meaningful.

            -1 if it a BAD COMMAND

        See Also
        --------
        zSetNSCSolve()
        """
        nscSolveData = -1
        cmd = "GetNSCSolve,{:d},{:d},{:d}".format(surfaceNumber,objectNumber,parameter)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if 'BAD COMMAND' not in rs:
            nscSolveData = tuple([float(e) if i in (3,4) else int(float(e))
                                 for i,e in enumerate(rs.split(","))])
        return nscSolveData

    def zGetOperand(self, row, column):
        """Returns the operand data from the Merit Function Editor

        Parameters
        ----------
        row : integer
            row operand number in the MFE
        column : integer
            column

        Returns
        -------
        operandData : integer/float/string
            opernadData's type depends on ``column`` argument if successful,
            else -1.

            Refer to the column-operand-data_ table for information on the
            types of ``operandData`` and ``column`` number

        Notes
        -----

        .. _column-operand-data:

        ::

            Table: Column and operand data types

            -----------------------------------------------
            column          operand data
            -----------------------------------------------
            1               operand type (string)
            2               int1 (integer)
            3               int2 (integer)
            4-7             data1-data4 (float)
            8               target (float)
            9               weight (float)
            10              value (float)
            11              percentage contribution (float)
            12-13           data5-data6 (float)

            end-of-table

        See Also
        --------
        zOptimize() :
            To update the merit function prior to calling ``zGetOperand()``,
            call ``zOptimize()`` with the number of cycles set to -1
        zSetOperand()
        """
        cmd = "GetOperand,{:d},{:d}".format(row, column)
        reply = self._sendDDEcommand(cmd)
        return _process_get_set_Operand(column, reply)

    def zGetPath(self):
        """Returns path-name-to-<data> folder and default lenses folder

        Parameters
        ----------
        None

        Returns
        -------
        pathToDataFolder : string
            full path to the <data> folder
        pathToDefaultLensFolder : string
            full path to the default folder for lenses
        """
        reply = str(self._sendDDEcommand('GetPath'))
        rs = str(reply.rstrip())
        return tuple(rs.split(','))

    def zGetPolState(self):
        """Returns the default polarization state set by the user

        Parameters
        ----------
        None

        Returns
        -------
        nlsPol : integer
            if ``nlsPol > 0``, then default polarization state is unpolarized
        Ex : float
            normalized electric field magnitude in x direction
        Ey : float
            normalized electric field magnitude in y direction
        Phax : float
            relative phase in x direction in degrees
        Phay : float
            relative phase in y direction in degrees

        Notes
        -----
        The quantity Ex*Ex + Ey*Ey should have a value of 1.0, although any
        values are accepted.

        See Also
        --------
        zSetPolState()
        """
        reply = self._sendDDEcommand("GetPolState")
        rs = reply.rsplit(",")
        polStateData = [int(float(elem)) if i==0 else float(elem)
                                       for i,elem in enumerate(rs[:-1])]
        return tuple(polStateData)

    def zGetPolTrace(self, waveNum, mode, surf, hx, hy, px, py, Ex, Ey, Phax, Phay):
        """Trace a single polarized ray through the lens system

        If ``Ex``, ``Ey``, ``Phax``, ``Phay`` are all zero, two orthogonal rays are
        traced, and the resulting transmitted intensity is averaged.

        Parameters
        ----------
        waveNum : integer
            wavelength number as in the wavelength data editor
        mode : integer (0/1)
            0 = real, 1 = paraxial
        surf : integer
            surface to trace the ray to. if -1, surf is the image plane.
        hx : float
            normalized field height along x axis
        hy : float
            normalized field height along y axis
        px : float
            normalized height in pupil coordinate along x axis
        py : float
            normalized height in pupil coordinate along y axis
        Ex : float
            normalized electric field magnitude in x direction
        Ey : float
            normalized electric field magnitude in y direction
        Phax : float
            relative phase in x direction in degrees
        Phay : float
            relative phase in y direction in degrees

        Returns
        -------
        rayPolTraceData : 8-tuple/ 2-tuple
            rayPolTraceData is a 8-tuple or 2-tuple (depending on polarized or
            unpolarized rays) containing the following elements:

            .. _returns-GetPolTrace:

            | For polarized rays:
            |    error       : 0, if the ray traced successfully;
            |                  +ve number indicates ray missed the surface
            |                  -ve number indicates ray total internal
            |                  reflected (TIR) at the surface given by the
            |                  absolute value of the errorCode number
            |    intensity   : the transmitted intensity of the ray. It is always
            |                  normalized to an input electric field intensity of
            |                  unity. The transmitted intensity accounts for
            |                  surface, thin film, and bulk absorption effects,
            |                  but does not consider whether or not the ray was
            |                  vignetted.
            |    Exr,Eyr,Ezr : real parts of the electric field components
            |    Exi,Eyi,Ezi : imaginary parts of the electric field components
            |
            | For unpolarized rays:
            |    error       : (see above)
            |    intensity   : (see above)

        Examples
        --------
        To trace the real unpolarized marginal ray to the image surface at
        wavelength 2, the function would be:

        >>> ln.zGetPolTrace(2, 0, -1, 0.0, 0.0, 0.0, 1.0, 0, 0, 0, 0)

        .. _notes-GetPolTrace:

        Notes
        -----
        1. The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
           values are accepted.
        2. There is an important exception to the above rule -- If Ex, Ey, Phax,
           Phay are all zero, Zemax will trace two orthogonal rays and the resul-
           ting transmitted intensity will be averaged.
        3. Always check to verify the ray data is valid (check the error) before
           using the rest of the data in the tuple.
        4. Use of ``zGetPolTrace()`` has significant overhead as only one ray per
           DDE call is traced. Please refer to the ZEMAX manual for more details.

        See Also
        --------
        zGetPolTraceDirect(), zGetTrace(), zGetTraceDirect()
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f}".format(px=px,py=py)
        args4 = "{Ex:1.4f},{Ey:1.4f}".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = "GetPolTrace," + args1 + args2 + args3 + args4 + args5
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rayPolTraceData = tuple([int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)])
        return rayPolTraceData

    def zGetPolTraceDirect(self, waveNum, mode, startSurf, stopSurf,
                           x, y, z, l, m, n, Ex, Ey, Phax, Phay):
        """Trace a single polarized ray using a more direct access to the Zemax
        ray tracing than ``zGetPolTrace()``

        If ``Ex``, ``Ey``, ``Phax``, ``Phay`` are all zero, Zemax will trace two
        orthogonal rays and the resulting transmitted intensity will be averaged.

        Parameters
        ----------
        waveNum : integer
            wavelength number as in the wavelength data editor
        mode : integer (0/1)
            0 = real, 1 = paraxial
        startSurf : integer
            surface to trace the ray from.
        stopSurf : integer
            last surface to trace the polarized ray to.
        x, y, z : floats
            coordinates of the ray at the starting surface
        l, m, n : floats
            the direction cosines to the entrance pupil aim point for the
            x-, y-, z- direction cosines respectively
        Ex : float
            normalized electric field magnitude in x direction
        Ey : float
            normalized electric field magnitude in y direction
        Phax : float
            relative phase in x direction in degrees
        Phay : float
            relative phase in y direction in degrees

        Returns
        -------
        rayPolTraceData : 8-tuple/ 2-tuple
            rayPolTraceData is the same data structure as that returned by
            ``zGetPolTrace()``. Refer to the description of the data structure
            returned by ``zGetPolTrace()`` (returns-GetPolTrace_) for details.

        Notes
        -----
        Refer to the notes (notes-GetPolTrace_) of ``zGetPolTrace()``

        See Also
        --------
        zGetPolTraceDirect(), zGetTrace(), zGetTraceDirect()
        """
        args0 = "{wN:d},{m:d},".format(wN=waveNum,m=mode)
        args1 = "{sa:d},{sd:d},".format(sa=startSurf,sd=stopSurf)
        args2 = "{x:1.20g},{y:1.20g},{y:1.20g},".format(x=x,y=y,z=z)
        args3 = "{l:1.20g},{m:1.20g},{n:1.20g},".format(l=l,m=m,n=n)
        args4 = "{Ex:1.4f},{Ey:1.4f}".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = "GetPolTraceDirect," + args0 + args1 + args2 + args3 + args4 + args5
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rayPolTraceData = tuple([int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)])
        return rayPolTraceData

    def zGetPupil(self):
        """Return the pupil data such as aperture type, ENPD, EXPD, etc.

        Parameters
        ----------
        None

        Returns
        -------
        aType : integer
            the system aperture de*ined as follows:

                * 0 = entrance pupil diameter
                * 1 = image space F/#
                * 2 = object space NA
                * 3 = float by stop
                * 4 = paraxial working F/#
                * 5 = object cone angle

        value : float
            the ``value`` is "stop surface semi-diameter" if
            ``aperture type == float by stop`` else ``value`` is the
            "sytem aperture"
        ENPD : float
            entrance pupil diameter (in lens units)
        ENPP : float
            entrance pupil position (in lens units)
        EXPD : float
            exit pupil diameter (in lens units)
        EXPP : float
            exit pupil position (in lens units)
        apodization_type : integer
            the apodization type is indicated as follows:

                * 0 = none
                * 1 = Gaussian
                * 2 = Tangential/Cosine cubed

        apodization_factor : float
            number shown on general data dialog box
        """
        reply = self._sendDDEcommand('GetPupil')
        rs = reply.split(',')
        pupilData = tuple([int(elem) if (i==0 or i==6)
                                 else float(elem) for i,elem in enumerate(rs)])
        return pupilData

    def zGetRefresh(self):
        """Copy lens data from the LDE into the Zemax server

        The lens is updated and Zemax re-computes all data.

        Parameters
        ----------
        None

        Returns
        -------
        status : integer (0, -1, or -998)
            0 if successful;
            -1 if Zemax could not copy the lens data from LDE to the server;
            -998 if the command times out (Note MZDDE returns -2)

        Notes
        -----
        If ``zGetRefresh()`` returns -1, no ray tracing can be performed.

        See Also
        --------
        zGetUpdate(), zPushLens()
        """
        reply = None
        reply = self._sendDDEcommand('GetRefresh')
        if reply:
            return int(reply) #Note: Zemax returns -1 if GetRefresh fails.
        else:
            return -998

    def zGetSag(self, surfaceNumber, x, y):
        """Return the sag of the surface at coordinates (x,y) in lens units

        Parameters
        ----------
        surfaceNumber : integer
            surface number
        x : float
            x coordinate in lens units
        y : float
            y coordinate in lens units

        Returns
        -------
        sag : float
            sag of the surface at (x,y) in lens units
        alternateSag : float
            alternate sag
        """
        cmd = "GetSag,{:d},{:1.20g},{:1.20g}".format(surfaceNumber,x,y)
        reply = self._sendDDEcommand(cmd)
        sagData = reply.rsplit(",")
        return (float(sagData[0]),float(sagData[1]))

    def zGetSequence(self):
        """Returns the sequence numbers of the lens in the server and in the LDE

        Parameters
        ----------
        None

        Returns
        -------
        seqNum_lenServ : float
            sequence number of lens in server
        seqNum_lenLDE : float
            sequence number of lens in LDE
        """
        reply = self._sendDDEcommand("GetSequence")
        seqNum = reply.rsplit(",")
        return (float(seqNum[0]),float(seqNum[1]))

    def zGetSerial(self):
        """Get the serial number of the running Zemax application

        Parameters
        ----------
        None

        Returns
        -------
        serial number : integer
            serial number
        """
        reply = self._sendDDEcommand('GetSerial')
        return int(reply.rstrip())

    def zGetSettingsData(self, tempFile, number):
        """Returns the settings data used by a window

        The data must have been previously stored by a call to
        ``zSetSettingsData()`` or by a previous execution of the client
        program.

        Parameters
        ----------
        tempfile : string
            the name of the output file passed by Zemax to the client. Zemax
            uses this name to identify for the window for which the
            ``zGetSettingsData()`` request is for.
        number : integer
            the data number used by the previous ``zSetSettingsData()`` call.
            Currently, only ``number = 0`` is supported.

        Returns
        -------
        settingsData : string
            data saved by a previous ``zSetSettingsData()`` call for the
            ``window`` and ``number``.

        See Also
        --------
        zSetSettingsData()
        """
        cmd = "GetSettingsData,{},{:d}".format(tempFile,number)
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())

    def zGetSolve(self, surfaceNumber, code):
        """Returns data about solves and/or pickups on the surface with number
        `surfaceNumber`.

        `zGetSolve(surfaceNumber, code)->solveData`

        Parameters
        ----------
        surfaceNumber : (integer) surface number
        code          : (integer) indicating which surface parameter the solve
                         is for, such as curvature, thickness, glass, conic,
                         semi-diameter, etc. (see the table below)
        Returns
        -------
        solveData     : a tuple, depending on the code value according to the
                        following table if successful, else -1.
        ------------------------------------------------------------------------
        code                    -  Returned data format
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

        Note
        ----
        The `solvetype` is an integer code, & the parameters have meanings
        that depend upon the solve type; see the chapter "SOLVES" in the Zemax
        manual for details.

        See also `zSetSolve`, `zGetNSCSolve`, `zSetNSCSolve`.
        """
        cmd = "GetSolve,{:d},{:d}".format(surfaceNumber,code)
        reply = self._sendDDEcommand(cmd)
        solveData = _process_get_set_Solve(reply)
        return solveData

    def zGetSurfaceData(self, surfaceNumber, code, arg2=None):
        """Gets surface data on a sequential lens surface.

        zGetSurfaceData(surfaceNum,code [, arg2])-> surfaceDatum

        Parameters
        ----------
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

        See also `zSetSurfaceData`, `zGetSurfaceParameter`.
        """
        if arg2 is None:
            cmd = "GetSurfaceData,{sN:d},{c:d}".format(sN=surfaceNumber,c=code)
        else:
            cmd = "GetSurfaceData,{sN:d},{c:d},{a:d}".format(sN=surfaceNumber,
                                                                 c=code,a=arg2)
        reply = self._sendDDEcommand(cmd)
        if code in (0,1,4,7,9):
            surfaceDatum = reply.rstrip()
        else:
            surfaceDatum = float(reply)
        return surfaceDatum

    def zGetSurfaceDLL(self, surfaceNumber):
        """Return the name of the DLL if the surface is a user defined type.

        zGetSurfaceDLL(surfaceNumber)->(dllName,surfaceName)

        Parameters
        ----------
        surfaceNumber: (integer) surface number of the user defined surface

        Returns
        -------
        Returns a tuble with the following elements
        dllName      : (string) The name of the defining DLL
        surfaceName  : (string) surface name displayed by the DLL in the surface
                       type column of the LDE.
        """
        cmd = "GetSurfaceDLL,{sN:d}".format(surfaceNumber)
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        return (rs[0],rs[1])

    def zGetSurfaceParameter(self, surfaceNumber, parameter):
        """Return the surface parameter data for the surface associated with the
        given surfaceNumber

        zGetSurfaceParameter(surfaceNumber,parameter)->parameterData

        Parameters
        ----------
        surfaceNumber  : (integer) surface number of the surface
        parameter      : (integer) parameter number ('Par' in LDE) being queried

        Returns
        --------
        parameterData  : (float) the parameter value

        Note
        ----
        To get thickness, radius, glass, semi-diameter, conic, etc, use
        `zGetSurfaceData`

        See also `zGetSurfaceData`, `zSetSurfaceParameter`.
        """
        cmd = "GetSurfaceParameter,{sN:d},{p:d}".format(sN=surfaceNumber,p=parameter)
        reply = self._sendDDEcommand(cmd)
        return float(reply)


    def zGetSystem(self):
        """Returns a number of general system data (General Lens Data)

        zGetSystem() -> systemData

        Returns
        -------
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

        Note
        ----
        The returned data structure is exactly similar to the data structure
        returned by the zSetSystem() method.

        See also zSetSystem, zGetFirst, zGetSystemProperty, zGetSystemAper, zGetAperture, zSetAperture

        Use `zGetFirst` to get first order lens data such as EFL, F/# etc.
        """
        sdt = _co.namedtuple('systemData' , ['numberOfSurfaces', 'unitCode',
                                             'stopSurfaceNum', 'nonAxialFlag',
                                             'rayAimingType', 'adjustIndexFlag',
                                             'temperature', 'pressure',
                                             'globalReferenceSurface'])
        reply = self._sendDDEcommand("GetSystem")
        rs = reply.split(',')
        systemData = sdt._make([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zGetSystemAper(self):
        """Gets system aperture data.

        zGetSystemAper()-> systemAperData

        Returns
        -------
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

        Note
        ----
        The returned tuple is the same as the returned tuple of zSetSystemAper()

        See also, `zGetSystem`, `zSetSystemAper`.
        """
        sad = _co.namedtuple('systemAper', ['apertureType', 'stopSurf', 'value'])
        reply = self._sendDDEcommand("GetSystemAper")
        rs = reply.split(',')
        systemAperData = sad._make([float(elem) if i==2 else int(float(elem))
                                    for i, elem in enumerate(rs)])
        return systemAperData

    def zGetSystemProperty(self, code):
        """Returns properties of the system, such as system aperture, field,
        wavelength, and other data, based on the integer `code` passed.

        zGetSystemProperty(code)-> sysPropData

        Parameters
        -----------
        code        : (integer) value that defines the specific system property
                      requested (see below).

        Returns
        -------
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

        See also zSetSystemProperty, zGetFirt
        """
        cmd = "GetSystemProperty,{c}".format(c=code)
        reply = self._sendDDEcommand(cmd)
        sysPropData = _process_get_set_SystemProperty(code,reply)
        return sysPropData

    def zGetTextFile(self, textFileName, analysisType, settingsFile=None, 
                     flag=0, timeout=None):
        """Generate a text file for any analysis that supports text output.

        Parameters
        ----------
        textFileName : string
            name of the file to be created including the full path and extension
        analysisType : string
            3 letter case-sensitive label that indicates the type of the analysis 
            to be performed. They are identical to the button codes. If no label 
            is provided or recognized, a standard raytrace will be generated
        settingsFile : string
            If ``settingsFile`` is valid, Zemax will use or save the settings 
            used to compute the text file, depending upon the value of the flag 
            parameter
        flag : integer (0/1/2)  
            0 = default settings used for the text;
            1 = settings provided in the settings file, if valid, else default;
            2 = settings provided in the settings file, if valid, will be used 
                and the settings box for the requested feature will be displayed. 
                After the user makes any changes to the settings the text will 
                then be generated using the new settings. Please see the manual 
                for more details
        timeout : integer, optional
            timeout in seconds (default=None, i.e. default timeout value)

        Returns
        -------
        retVal : integer
            0 = success;
            -1 = text file could not be saved (Zemax may not have received
                 a full path name or extention);
            -998 = command timed out

        Notes
        -----
        No matter what the flag value is, if a valid file name is provided for 
        ``settingsFile``, the settings used will be written to the settings
        file, overwriting any data in the file.

        See Also
        -------- 
        zGetMetaFile(), zOpenWindow()
        """
        retVal = -1
        if settingsFile:
            settingsFile = settingsFile
        else:
            settingsFile = ''
        #Check if the file path is valid and has extension
        if _os.path.isabs(textFileName) and _os.path.splitext(textFileName)[1]!='':
            cmd = 'GetTextFile,"{tF}",{aT},"{sF}",{fl:d}'.format(tF=textFileName,
                                    aT=analysisType,sF=settingsFile,fl=flag)
            reply = self._sendDDEcommand(cmd, timeout)
            if 'OK' in reply.split():
                retVal = 0
        return retVal

    def zGetTol(self,operandNumber):
        """Returns the tolerance data.

        zGetTol(operandNumber)->toleranceData

        Parameters
        ---------
        operandNumber : 0 or the tolerance operand number (row number in the
                        tolerance editor, when greater than 0)
        Returns
        ------
        toleranceData. It is a number or a 6-tuple, depending upon `operandNumber`
        as follows:
          if operandNumber == 0, toleranceData = number where `number` is the
          number of tolerance operands defined.
          if operandNumber > 0, toleranceData = (tolType, int1, int2, min, max, int3)

        See also `zSetTol`, `zSetTolRow`
        """
        reply = self._sendDDEcommand("GetTol,{:d}".format(operandNumber))
        if operandNumber == 0:
            toleranceData = int(float(reply.rstrip()))
            if toleranceData == 1:
                reply = self._sendDDEcommand("GetTol,1")
                tolType = reply.rsplit(",")[0]
                if tolType == 'TOFF': # the tol editor is actually empty
                    toleranceData = 0
        else:
            toleranceData = _process_get_set_Tol(operandNumber,reply)
        return toleranceData

    def zGetTrace(self, waveNum, mode, surf, hx, hy, px, py):
        """Trace a (single) ray through the current lens in the ZEMAX DDE server.

        zGetTrace(waveNum,mode,surf,hx,hy,px,py[,retNamedTuple]) -> rayTraceData

        Parameters
        ----------
        waveNum : wavelength number as in the wavelength data editor
        mode    : 0 = real, 1 = paraxial
        surf    : surface to trace the ray to. Usually, the ray data is only
                  needed at the image surface; setting the surface number to
                  -1 will yield data at the image surface.
        hx      : normalized field height along x axis
        hy      : normalized field height along y axis
        px      : normalized height in pupil coordinate along x axis
        py      : normalized height in pupil coordinate along y axis

        Returns
        -------
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

        Example:
        -------
        To trace the real chief ray to surface 5 for wavelength 3, use
        rayTraceData = zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

        OR

        (errorCode,vigCode,x,y,z,l,m,n,l2,m2,n2,intensity) = \
                                          zGetTrace(3,0,5,0.0,1.0,0.0,0.0)

        Note
        ----
        1. Always check to verify the ray data is valid  (errorCode) before using
           the rest of the string!
        2. Use of zGetTrace() has significant overhead as only one ray per DDE call
           is traced. Please refer to the ZEMAX manual for more details. Also, if a
           large number of rays are to be traced, see the section "Tracing large
           number of rays" in the ZEMAX manual.

        See also `zGetTraceDirect`, `zGetPolTrace`, `zGetPolTraceDirect`
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f}".format(px=px,py=py)
        cmd = "GetTrace," + args1 + args2 + args3
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rayData = [int(elem) if (i==0 or i==1)
                                  else float(elem) for i,elem in enumerate(rs)]
        rtd = _co.namedtuple('rayTraceData', ['errCode', 'vigCode',
                                              'x', 'y', 'z',
                                              'dcos_l', 'dcos_m', 'dcos_n',
                                              'dnorm_l2', 'dnorm_m2', 'dnorm_n2',
                                              'intensity'])
        rayTraceData = rtd._make(rayData)
        return rayTraceData

    def zGetTraceDirect(self, waveNum, mode, startSurf, stopSurf, x, y, z, l, m, n):
        """Trace a (single) ray through the current lens in the ZEMAX DDE server
        while providing a more direct access to the ZEMAX ray tracing engine than
        zGetTrace.

        zGetTraceDirect(waveNum,mode,startSurf,stopSurf,x,y,z,l,m,n)->rayTraceData

        Parameters
        ---------
        waveNum  : wavelength number as in the wavelength data editor
        mode     :  0 = real, 1 = paraxial
        startSurf: starting surface of the ray
        stopSurf : stopping surface of the ray
        x,y,z,   : coordinates of the ray at the starting surface
        l,m,n    : the direction cosines to the entrance pupil aim point for
                   the x-, y-, z- direction cosines respectively

        Returns
        -------
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
        ------
        Normally, rays are defined by the normalized field and pupil coordinates
        hx, hy, px, and py. ZEMAX takes these normalized coordinates and computes
        he object coordinates (x, y, and z) and the direction cosines to the
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
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rtd = _co.namedtuple('rayTraceData', ['errCode', 'vigCode',
                                              'x', 'y', 'z',
                                              'dcos_l', 'dcos_m', 'dcos_n',
                                              'dnorm_l2', 'dnorm_m2', 'dnorm_n2',
                                              'intensity'])
        rayTraceData = rtd._make([int(elem) if (i==0 or i==1)
                                  else float(elem) for i,elem in enumerate(rs)])
        return rayTraceData

    def zGetUDOSystem(self, bufferCode):
        """Load a particular lens from the optimization function memory into the
        ZEMAX server's memory. This will cause ZEMAX to retrieve the correct lens
        from system memory, and all subsequent DDE calls will be for actions
        (such as ray tracing) on this lens.The only time this item name should be
        used is when implementing a User Defined Operand, or UDO.

        zGetUDOSystem(bufferCode)->

        Parameters
        ----------
        bufferCode: (integer) The buffercode is an integer value provided by
                    ZEMAX to the client that uniquely identifies the correct lens.

        Returns
        ------
          ?

        Note
        ----
        Once the data is computed, up to 1001 values may be sent back to
        the server, and ultimately to the optimizer within ZEMAX, with the
        zSetUDOItem command.

        See also zSetUDOItem.
        """
        cmd = "GetUDOSystem,{:d}".format(bufferCode)
        reply = self._sendDDEcommand(cmd)
        return _regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zGetUpdate(self):
        """Update the lens, which means Zemax recomputes all pupil positions,
        solves, and index data.

        zGetUpdate()->status

        Parameters
        ----------
        None

        Returns
        -------
        status :   0 = Zemax successfully updated the lens
                  -1 = No raytrace performed
                -998 = Command timed out

        To update the merit function, use the zOptimize item with the number
        of cycles set to -1.

        See also zGetRefresh, zOptimize, zPushLens
        """
        status,ret = -998, None
        ret = self._sendDDEcommand("GetUpdate")
        if ret != None:
            status = int(ret)  #Note: Zemax returns -1 if GetUpdate fails.
        return status

    def zGetVersion(self):
        """Get the current version of ZEMAX which is running.

        zGetVersion() -> version (integer, generally 5 digit)

        """
        return int(self._sendDDEcommand("GetVersion"))

    def zGetWave(self, n):
        """Extract wavelength data from ZEMAX DDE server.

        There are 2 ways of using this function:
            zGetWave(0)-> waveData
              OR
            zGetWave(wavelengthNumber)-> waveData

        Returns
        -------
        if n==0: waveData is a tuple containing the following:
            primary : number indicating the primary wavelength (integer)
            number  : number_of_wavelengths currently defined (integer).
        elif 0 < n <= number_of_wavelengths: waveData consists of
            wavelength : value of the specific wavelength (floating point)
            weight     : weight of the specific wavelength (floating point)

        Note
        ----
        The returned tuple is exactly same in structure and contents to that
        returned by zSetWave().

        See also zSetWave(), zSetWaveTuple(), zGetWaveTuple().
        """
        reply = self._sendDDEcommand('GetWave,' + str(n))
        rs = reply.split(',')
        if n:
            wtd = _co.namedtuple('waveData', ['wavelength', 'weight'])
            waveData = wtd._make([float(ele) for ele in rs])
        else:
            wtd = _co.namedtuple('waveData', ['primaryWavelengthNum',
                                              'numberOfWavelengths'])
            waveData = wtd._make([int(ele) for ele in rs])
        return waveData

    def zGetWaveTuple(self):
        """Gets data on all defined wavelengths from the ZEMAX DDE server. This
        function is similar to "zGetWaveDataMatrix()" in MZDDE toolbox.

        zDetWaveTuple() -> ((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))

        Parameters
        ---------
        None

        Returns
        ------
        waveDataTuple: wave data tuple is a 2D tuple with the first
                       dimension (first subtuple) containing the wavelengths and
                       the second dimension containing the weights like so:
                       ((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))

        See also zSetWaveTuple(), zGetWave(), zSetWave()
        """
        waveCount = self.zGetWave(0)[1]
        waveData = [[],[]]
        wdt = _co.namedtuple('waveDataTuple', ['wavelengths', 'weights'])
        for i in range(waveCount):
            cmd = "GetWave,{wC:d}".format(wC=i+1)
            reply = self._sendDDEcommand(cmd)
            rs = reply.split(',')
            waveData[0].append(float(rs[0])) # store the wavelength
            waveData[1].append(float(rs[1])) # store the weight
        waveDataTuple = wdt(tuple(waveData[0]),tuple(waveData[1]))
        return waveDataTuple

    def zHammer(self, numOfCycles, algorithm, timeout=60):
        """Calls the Hammer optimizer.

        zHammer(numOfCycles,algorithm)->finalMeritFn

        Parameters
        ---------
        numOfCycles  : (integer) the number of cycles to run
                     if numOfCycles < 1, zHammer updates all operands in the
                     merit function and returns the current merit function,
                     and no optimization is performed.
        algorithm    : 0 = Damped Least Squares
                     1 = Orthogonal descent
        timeout      : (integer) timeout in seconds (default=1min)

        Returns
        -------
        finalMeritFn : (float) the final merit function.

        Note
        ----
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
        cmd = "Hammer,{:1.2g},{:d}".format(numOfCycles, algorithm)
        reply = self._sendDDEcommand(cmd, timeout)
        return float(reply.rstrip())

    def zImportExtraData(self, surfaceNumber, fileName):
        """Imports extra data and grid surface data values into an existing sur-
        face.

        zImportExtraData(surfaceNumber,fileName)->?

        Parameters
        ----------
        surfaceNumber : (integer) surface number
        fileName      : (string) file name (of an ASCII file)

        Returns
        -------
         ?

        Note
        -----
        The ASCII file is a single column of free-format numbers, with a
        .DAT extension.
        """
        cmd = "ImportExtraData,{:d},{}".format(surfaceNumber,fileName)
        reply = self._sendDDEcommand(cmd)
        return reply.rstrip()
        # !!! FIX determine what is the currect return


    def zInsertConfig(self, configNumber):
        """Insert a new configuration (column) in the multi-configuration editor.
        The new configuration will be placed at the location (column) indicated
        by the parameter `configNumber`.

        zInsertConfig(configNumber)->configNumberRet

        Parameters
        ---------
        configNumber  : (integer) the configuration (column) number to insert.

        Returns
        -------
        configNumberRet : (integer) the column number of the configuration that
                          is inserted at configNumber.

        Note
        ----
        1. The configNumber returned (configNumberRet) is generally different
           from the number in the input configNumber.
        2. Use zInsertMCO() to insert a new multi-configuration operand in the
           multi-configuration editor.
        3. Use zSetConfig() to switch the current configuration number

        See also zDeleteConfig().
        """
        return int(self._sendDDEcommand("InsertConfig,{:d}".format(configNumber)))

    def zInsertMCO(self,operandNumber):
        """Insert a new multi-configuration operand (row) in the multi-configuration
        editor.

        zInsertMCO(operandNumber)-> retValue

        Parameters
        ---------
        operandNumber : (integer) between 1 and the current number of operands
                        plus 1, inclusive.

        Returns
        -------
        retValue      : new number of operands (rows).

        See also zDeleteMCO. Use zInsertConfig(), to insert a new configuration (row).
        """
        return int(self._sendDDEcommand("InsertMCO,{:d}".format(operandNumber)))

    def zInsertMFO(self,operandNumber):
        """Insert a new optimization operand (row) in the merit function editor.

        zInsertMFO(operandNumber)->retValue

        Parameters
        ----------
        operandNumber : (integer) between 1 and the current number of operands
                        plus 1, inclusive.

        Returns
        -------
        retValue      : new number of operands (rows).

        See also zDeleteMFO. Generally, you may want to use zSetOperand() afterwards.
        """
        return int(self._sendDDEcommand("InsertMFO,{:d}".format(operandNumber)))

    def zInsertObject(self, surfaceNumber, objectNumber):
        """
        Insert a new NSC object at the location indicated by the parameters
        `surfaceNumber` and `objectNumber`.

        zInsertObject(surfaceNumber,objectNumber)->status

        Parameters
        ---------
        surfaceNumber : (integer) surface number of the NSC group. Use 1 if
                        the program mode is Non-Sequential.
        objectNumber  : object number

        Returns
        -------
        status        : 0 if successful, -1 if failed.

        See also zSetNSCObjectData to define data for the new surface and the
        zDeleteObject function.
        """
        cmd = "InsertObject,{:d},{:d}".format(surfaceNumber,objectNumber)
        reply = self._sendDDEcommand(cmd)
        if reply.rstrip() == 'BAD COMMAND':
            return -1
        else:
            return int(reply.rstrip())

    def zInsertSurface(self,surfNum):
        """Insert a lens surface in the ZEMAX DDE server.

        The new surface will be placed at the location indicated by the parameter
        `surfNum`.

        zInsertSruface(surfNum)-> retVal

        Parameters
        ---------
        surfNum  : location where to insert the surface

        Returns
        ------
        retVal : 0 if success

        See also zSetSurfaceData() to define data for the new surface and the
        zDeleteSurface() functions.
        """
        return int(self._sendDDEcommand("InsertSurface,"+str(surfNum)))

    def zLoadDetector(self, surfaceNumber, objectNumber, fileName):
        """Loads the data saved in a file to an NSC Detector Rectangle, Detector
        Color, Detector Polar, or Detector Volume object.

        zLoadDetector(surfaceNumber, objectNumber, fileName)->status

        Parameters
        ----------
        surfNumber   : (integer) surface number of the non-sequential group.
                     Use 1 if the program mode is Non-Sequential.
        objectNumber : (integer) object number of the detector object
        fileName     : (string) The filename may include the full path, if no
                     path is provided the path of the current lens file is
                     used. The extension should be DDR, DDC, DDP, or DDV for
                     Detector Rectangle, Color, Polar, and Volume objects,
                     respectively.
        Returns
        -------
        status       : 0 if load was successful
                       Error code (such as -1,-2) if failed.
        """
        isRightExt = _os.path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not _os.path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        isFile = _os.path.isfile(fileName)  # check if file exist
        if isRightExt and isFile:
            cmd = ("LoadDetector,{:d},{:d},{}"
                   .format(surfaceNumber,objectNumber,fileName))
            reply = self._sendDDEcommand(cmd)
            return _regressLiteralType(reply.rstrip())
        else:
            return -1

    def zLoadFile(self, fileName, append=None):
        """Loads a ZEMAX file into the server.

        zLoadFile(fileName[,append]) -> retVal

        Parameters
        ----------
        filename: full path of the ZEMAX file to be loaded. For example:
                  "C:\ZEMAX\Samples\cooke.zmx"
        append  : (optional, integer) If a non-zero value of append is passed,
                  then the new file is appended to the current file starting
                  at the surface number defined by the value appended.

        Returns
        -------
        retVal:    0: file successfully loaded
                -999: file could not be loaded (check if the file really exists, or
                      check the path.
                -998: the command timed out
               other: the upload failed.

        See also zSaveFile, zGetPath, zPushLens, zuiLoadFile
        """
        reply = None
        isAbsPath = _os.path.isabs(fileName)
        isRightExt = _os.path.splitext(fileName)[1] in ('.zmx','.ZMX')
        isFile = _os.path.isfile(fileName)
        if isAbsPath and isRightExt and isFile:
            if append:
                cmd = "LoadFile,{},{}".format(fileName,append)
            else:
                cmd = "LoadFile,{}".format(fileName)
            reply = self._sendDDEcommand(cmd)
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

        Parameters
        ----------
        fileName  : (string) name of the merit function file with full path and
                     extension.
        Returns
        -------
        meritData : If the loading is successful, meritData is a 2-tuple
                    containing the following elements:
                    number : (integer) number of operands in the merit function
                    merit  : (float) merit value of the merit function.

                     If meritData = -999, file could not be loaded (check if the
                    file really exists, or check the path.
        Note
        -----
        1. If the merit function value is 9.00e+009, the merit function cannot
           be evaluated.
        2. Loading a merit function file does not change the data displayed in
           the LDE; the server process has a separate copy of the lens data.

        See also: zOptimize, zSaveMerit
        """
        isAbsPath = _os.path.isabs(fileName)
        isRightExt = _os.path.splitext(fileName)[1] in ('.mf','.MF','.zmx','.ZMX')
        isFile = _os.path.isfile(fileName)
        if isAbsPath and isRightExt and isFile:
            reply = self._sendDDEcommand('LoadMerit,'+fileName)
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

        Parameters
        ----------
        fileName : (string) file name of the tolerance file. If no path is
                   provided in the filename, the <data>\Tolerance folder is
                   assumed.
        Returns
        -------
        numTolOperands : number of tolerance operands loaded.
                         -999 if file doesnot exist
        """
        if _os.path.isabs(fileName): # full path is provided
            fullFilePathName = fileName
        else:                    # full path not provided
            fullFilePathName = self.zGetPath()[0] + "\\Tolerance\\" + fileName
        if _os.path.isfile(fullFilePathName):
            cmd = "LoadTolerance,{}".format(fileName)
            reply = self._sendDDEcommand(cmd)
            return int(float(reply.rstrip()))
        else:
            return -999

    def zMakeGraphicWindow(self, fileName, moduleName, winTitle, textFlag, settingsData=None):
        """Notifies ZEMAX that graphic data has been written to a file and may now
        be displayed as a ZEMAX child window. The primary purpose of this item is
        to implement user defined features in a client application, that look and
        act like native ZEMAX features.

        There are two ways of using this command:

        zMakeGraphicWindow(fileName,moduleName,winTitle,textFlag,settingsData)->retString

          OR

        zSetSettingsData(0,settingsData)
        zMakeGraphicWindow(self,fileName,moduleName,winTitle,textFlag)->retString

        Parameters
        ----------
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

        Returns
        -------
        None

        Examples
        --------
        A sample item string might look like the following:

        >>> ln.zMakeGraphicWindow('C:\TEMP\ZGF001.TMP',
                                  'C:\ZEMAX\FEATURES\CLIENT.EXE',
                                  'ClientWindow', 1, "0 1 2 12.55")

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
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())
        # FIX !!! What is the appropriate reply?

    def zMakeTextWindow(self, fileName, moduleName, winTitle, settingsData=None):
        """Notifies Zemax that text data has been written to a file and may now
        be displayed as a Zemax child window. The primary purpose of this item
        is to implement user defined features in a client application, that look
        and act like native Zemax features.

        There are two ways of using this command:

        zMakeTextWindow(fileName,moduleName,winTitle,settingsData)->retString

          OR

        zSetSettingsData(0,settingsData)
        zMakeTextWindow(self,fileName,moduleName,winTitle)->retString

        Parameters
        ----------
        fileName     : the full path and file name to the temporary file that
                       holds the text data. This must be the same name as
                       passed to the client executable in the command line
                       arguments, if any.
        moduleName   : the full path and executable name of the client program
                       that created the text data.
        winTitle     : the string which defines the title Zemax should place in
                       the top bar of the window.
        settingsData : The settings data is a string of values delimited by
                       spaces (not commas) which are used by the client to define
                       how the data was generated. These values are only used
                       by the client, not by Zemax. The settings data string
                       holds the options and data that would normally appear in
                       a Zemax "settings" style dialog box. The settings data
                       should be used to recreate the data if required. Because
                       the total length of a data item cannot exceed 255
                       characters, the function zSetSettingsData() may be used
                       prior to the call to zMakeTextWindow() to specify the
                       settings data string rather than including the data as
                       part of zMakeTextWindow(). See "How Zemax calls the
                       client" in the manual for more details on the settings
                       data.

        A sample item string might look like the following:

        zMakeTextWindow('C:\TEMP\ZGF002.TMP','C:\ZEMAX\FEATURES\CLIENT.EXE',
                           'ClientWindow',"6 5 4 12.55")

        This call indicates that Zemax should open a text window, display the
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
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())
        # FIX !!! What is the appropriate reply?

    def zModifySettings(self, fileName, mType, value):
        """change specific options in Zemax settings files (.CFG)

        Settings files are used by ``zMakeTextWindow()`` & ``zMakeGraphicWindow()``.
        The modified settings file is written back to the original settings fileName.

        Parameters
        ----------
        fileName : string
            full name of the settings file, including the path & extension
        mType : string
            a mnemonic that indicates which setting within the file is to be 
            modified. See the ZPL macro command "MODIFYSETTINGS" in the Zemax 
            manual for a complete list of the ``mType`` codes
        value : string/integer
            the new data for the specified setting

        Returns
        -------
        status : integer
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

        Examples
        --------
        >>> ln.zModifySettings("C:\MyPOP.CFG", "POP_BEAMTYPE", 2)
        """
        if isinstance(value, str):
            cmd = "ModifySettings,{},{},{}".format(fileName,mType,value)
        else:
            cmd = "ModifySettings,{},{},{:1.20g}".format(fileName,mType,value)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply.rstrip()))

    def zNewLens(self):
        """erases the current lens

        The "minimum" lens that remains is identical to the LDE when "File>>New" 
        is selected. No prompt to save the existing lens is given

        Parameters
        ---------- 
        None

        Returns
        ------- 
        status : integer
            0 = successful
        """
        return int(self._sendDDEcommand('NewLens'))

    def zNSCCoherentData(self,surfaceNumber,objectNumDisDetectr,pixel,dataType):
        """Return data from an NSC detector (Non-sequential coherent data)

        zNSCCoherentData(surfaceNumber,objectNumDisDetectr,pixel,dataType)->

        Parameters
        ----------
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
        reply = self._sendDDEcommand(cmd)
        return float(reply.rstrip())

    def zNSCDetectorData(self,surfaceNumber,objectNumDisDetectr,pixel,dataType):
        """Return data from an NSC detector (Non-sequential incoherent intensity
        data)

        zNSCDetectorData(surfaceNumber,objectNumDisDetectr,pixel,dataType)->

        Parameters
        ----------
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
        reply = self._sendDDEcommand(cmd)
        return float(reply.rstrip())

    def zNSCLightningTrace(self, surfNumber, source, raySampling, edgeSampling, timeout=60):
        """Traces rays from one or all NSC sources using Lighting Trace.

        zNSCLightningTrace(surfNumber, source, raySampling, edgeSampling) ->

        Parameters
        ---------
        surfNumber   : (integer) surface number, use 1 for pure NSC mode
        source       : (integer) object number of the desired source. If source
                      is zero, all sources will be traced.
        raySampling  : resolution of the LightningTrace mesh with valid values
                     between 0 (= "Low (1X)") and 5 (= "1024X").
        edgeSampling : resolution used in refining the LightningTrace mesh near
                     the edges of objects, with valid values between 0 ("Low (1X)")
                     and 4 ("256X").
        timeout      : timeout value in seconds. Default=60sec

        Note: `zNSCLightningTrac`e always updates the lens before executing a
        LightningTrace to make certain all objects are correctly loaded and
        updated.
        """
        cmd = ("NSCLightningTrace,{:d},{:d},{:d},{:d}"
               .format(surfNumber, source, raySampling, edgeSampling))
        reply = self._sendDDEcommand(cmd, timeout)
        if 'OK' in reply.split():
            return 0
        elif 'BAD COMMAND' in reply.rstrip():
            return -1
        else:
            return int(float(reply.rstrip()))  # return the error code sent by zemax.

    def zNSCTrace(self, surfNum, objNumSrc, split=0, scatter=0, usePolar=0,
                  ignoreErrors=0, randomSeed=0, save=0, saveFilename=None,
                  oFilter=None, timeout=60):
        """Traces rays from one or all NSC sources with various optional arguments.
        zNSCTrace() always updates the lens before tracing rays to make certain all
        objects are correctly loaded and updated.

        Possible ways of using this function:

        zNSCTrace(surfNum,objNumSrc,[split,scatter,usePolar,ignoreErrors,
                  randomSeed,save=0])->traceResult

          OR

        zNSCTrace(surfNum,objNumSrc,[split,scatter,usePolar,ignoreErrors,
                  randomSeed,]save,saveFilename,oFilter)->traceResult

        Parameters
        ----------
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
        timeout      : timeout in seconds (default = 60 seconds)

        Returns
        -------
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
            isAbsPath = _os.path.isabs(saveFilename)
            isRightExt = _os.path.splitext(saveFilename)[1] in ('.ZRD',)
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
        reply = self._sendDDEcommand(cmd, timeout)
        if 'OK' in reply.split():
            return 0
        elif 'BAD COMMAND' in reply.rstrip():
            return -1
        else:
            return int(float(reply.rstrip()))  # return the error code sent by zemax.

    def zOpenWindow(self, analysisType, zplMacro=False, timeout=None):
        """Open a new analysis window on the main ZEMAX screen.

        `zOpenWindow(analysisType)->status`

        Parameters
        ---------
        analysisType : (string) 3 character case-sensitive label that indicates
                      the type of analysis to be performed. The 3 letter
                      labels are identical to those used for the button bar
                      in ZEMAX. You can see a list of the button codes by
                      importing zemaxbuttons module and calling the showZButtons
                      function in an interactive shell, for example:
                          `import zemaxbuttons as zb`
                          `zb.showZButtonList()`
        zplMacro      : (bool) True if the analysisType code is the first 3-letters
                      of a ZPL macro name, else False (default). Please see
                      the Note below
        timeout       : timeout value in seconds. Default=None

        Returns
        -------
        status  : 0 = successful, -1 = incorrect analysis code, -999 = 'FAIL'

        Note
        ----
        The function checks if the analysisType code is a valid code or not in
        order to prevent the calling program to get stalled. However, it doesn't
        check the analysisType code validity if it is a ZPL macro code.
        If the analysisType is ZPL macro, please make sure that the macro exist
        in the <data>/Macros folder. It is recommended to use zExecuteZPLMacro()
        if you are trying to execute a ZPL macro.

        See also `zGetMetaFile`, `zExecuteZPLMacro`
        """
        if zb.isZButtonCode(analysisType) ^ zplMacro:  # XOR operation
            reply = self._sendDDEcommand("OpenWindow,{}".format(analysisType), timeout)
            if 'OK' in reply.split():
                return 0
            elif 'FAIL' in reply.split():
                return -999
            else:
                return int(float(reply.rstrip()))  # error code from Zemax
        else:
            return -1 # Incorrect analysisType code

    def zOperandValue(self, operandType, *values):
        """Returns the value of any optimization operand, even if the operand is
        not currently in the merit function.

        `zOperandValue(operandType,*values)->operandValue`

        Parameters
        ----------
        operandType  : a valid optimization operand
        *values      : a sequence of arguments. Possible arguments include
                       int1, int2, data1, data2, data3, data4, data5, data6
        Returns
        -------
        operandValue : (float) the value
        """
        if zo.isZOperand(operandType,1) and (0 < len(values) < 9):
            valList = [str(int(elem)) if i in (0,1) else str(float(elem))
                       for i,elem in enumerate(values)]
            arguments = ",".join(valList)
            cmd = "OperandValue," + operandType + "," + arguments
            reply = self._sendDDEcommand(cmd)
            return float(reply.rstrip())
        else:
            return -1

    def zOptimize(self, numOfCycles=0, algorithm=0, timeout=None):
        """Calls the Zemax Damped Least Squares/ Orthogonal Descent optimizer.

        `zOptimize(numOfCycles,algorithm)->finalMeritFn`

        Parameters
        ----------
        numOfCycles  : (integer) the number of cycles to run
                       if numOfCycles == 0 (default), optimization runs in automatic mode.
                       if numOfCycles < 0, zOptimize updates all operands in the
                       merit function and returns the current merit function,
                       and no optimization is performed.
        algorithm    : 0 = Damped Least Squares
                       1 = Orthogonal descent
        timeout      : timeout value in seconds

        Returns
        -------
        finalMeritFn : (float) the final merit function.

        Note
        ----
        1. If the merit function value returned is 9.0E+009, the optimization
           failed, usually because the lens or merit function could not be
           evaluated.
        2. The number of cycles should be kept small enough to allow the algorithm
           to complete and return before the DDE communication times out, or
           an error will occur. One possible way to achieve high number of
           cycles could be to call zOptimize multiple times in a loop, each time
           comparing the returned merit function with few of the previously
           returned (& stored) merit functions to determine if an optimum has
           been attained. For an example implementation see `zOptimize2()`

        See also `zHammer`, `zLoadMerit`, `zsaveMerit`, `zOptimize2`
        """
        cmd = "Optimize,{:1.2g},{:d}".format(numOfCycles,algorithm)
        reply = self._sendDDEcommand(cmd, timeout)
        return float(reply.rstrip())

    def zPushLens(self, update=None, timeout=None):
        """Copy lens in the ZEMAX DDE server into the Lens Data Editor (LDE).

        `zPushLens([update, timeout]) -> retVal`

        Parameters
        ---------
        updateFlag (optional): if 0 or omitted, the open windows are not updated.
                               if 1, then all open analysis windows are updated.
        timeout (optional)   : if a timeout in seconds in passed, the client will
                               wait till the timeout before returning a timeout
                               error. If no timeout is passed, the default timeout
                               of 3 seconds is used.

        Returns
        -------
        retVal:
                0: lens successfully pushed into the LDE.
             -999: the lens could not be pushed into the LDE. (check zPushLensPermission)
             -998: the command timed out
            other: the update failed.

        Note
        -----
        This operation requires the permission of the user running the ZEMAX program.
        The proper use of `zPushLens` is to first call `zPushLensPermission`.

        See also `zPushLensPermission`, `zLoadFile`, `zGetUpdate`, `zGetPath`,
        `zGetRefresh`, `zSaveFile`.
        """
        reply = None
        if update == 1:
            reply = self._sendDDEcommand('PushLens,1', timeout)
        elif update == 0 or update is None:
            reply = self._sendDDEcommand('PushLens', timeout)
        else:
            raise ValueError('Invalid value for flag')
        if reply:
            return int(reply)   # Note: Zemax returns -999 if push lens fails
        else:
            return -998         # if timeout reached (assumption!!)

    def zPushLensPermission(self):
        """Establish if ZEMAX extensions are allowed to push lenses in the LDE.

        `zPushLensPermission() -> status`

        Parameters
        ---------
        None

        Return
        ------
        status:
            1: ZEMAX is set to accept PushLens commands
            0: Extensions are not allowed to use PushLens

        For more details, please refer to the ZEMAX manual.

        See also `zPushLens`, `zGetRefresh`
        """
        status = None
        status = self._sendDDEcommand('PushLensPermission')
        return int(status)

    def zQuickFocus(self,mode=0,centroid=0):
        """Performs a quick best focus adjustment for the optical system by adjusting
        the back focal distance for best focus. The "best" focus is chosen as a wave-
        length weighted average over all fields. It adjusts the thickness of the
        surface prior to the image surface.

        `zQuickFocus([mode,centroid]) -> retVal`

        Parameters
        ----------
        mode:
            0: RMS spot size (default)
            1: spot x
            2: spot y
            3: wavefront OPD
        centroid: to specify RMS reference
            0: RMS referenced to the chief ray (default)
            1: RMS referenced to image centroid

        Returns
        -------
        retVal: 0 for success.
        """
        retVal = -1
        cmd = "QuickFocus,{mode:d},{cent:d}".format(mode=mode,cent=centroid)
        reply = self._sendDDEcommand(cmd)
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zReleaseWindow(self, tempFileName):
        """Release locked window/menu mar.

        zReleaseWindow(tempFileName)->status

        Parameters
        ----------
        tempFileName  : (string)

        Returns
        -------
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
        reply = self._sendDDEcommand("ReleaseWindow,{}".format(tempFileName))
        return int(float(reply.rstrip()))

    def zRemoveVariables(self):
        """Sets all currently defined variables to fixed status.

        zRemoveVariables()->status

        Parameters
        ----------
        None

        Returns
        -------
        status : 0 = successful.
        """
        reply = self._sendDDEcommand('RemoveVariables')
        if 'OK' in reply.split():
            return 0
        else:
            return -1

    def zSaveDetector(self, surfaceNumber, objectNumber, fileName):
        """Saves the data currently on an NSC Detector Rectangle, Detector Color,
        Detector Polar, or Detector Volume object to a file.

        zSaveDetector(surfaceNumber, objectNumber, fileName)->status

        Parameters
        ----------
        surfNumber   : (integer) surface number of the non-sequential group.
                       Use 1 if the program mode is Non-Sequential.
        objectNumber : (integer) object number of the detector object
        fileName     : (string) The filename may include the full path, if no
                       path is provided the path of the current lens file is
                       used. The extension should be DDR, DDC, DDP, or DDV for
                       Detector Rectangle, Color, Polar, and Volume objects,
                       respectively.
        Returns
        -------
        status       : 0 if save was successful
                       Error code (such as -1,-2) if failed.
        """
        isRightExt = _os.path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not _os.path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        if isRightExt:
            cmd = ("SaveDetector,{:d},{:d},{}"
                   .format(surfaceNumber,objectNumber,fileName))
            reply = self._sendDDEcommand(cmd)
            return _regressLiteralType(reply.rstrip())
        else:
            return -1

    def zSaveFile(self, fileName):
        """Saves the lens currently loaded in the server to a ZEMAX file.

        zSaveFile(fileName)-> status

        Parameters
        ----------
        fileName : (string) file name, including full path with extension.

        Returns
        -------
        status   :    0 = Zemax successfully saved the lens file & updated the
                          newly saved lens
                   -999 = Zemax couldn't save the file
                     -1 = Incorrect file name
                   Any other value = update failed.

        See also zGetPath, zGetRefresh, zLoadFile, zPushLens.
        """
        isAbsPath = _os.path.isabs(fileName)
        isRightExt = _os.path.splitext(fileName)[1] in ('.zmx','.ZMX')
        if isAbsPath and isRightExt:
            cmd = "SaveFile,{}".format(fileName)
            reply = self._sendDDEcommand(cmd)
            return int(float(reply.rstrip()))
        else:
            return -1

    def zSaveMerit(self, fileName):
        """Saves the current merit function to a ZEMAX .MF file.

        zSaveMerit(fileName)->meritData

        Parameters
        ----------
        fileName  : (string) name of the merit function file with full path and
                     extension.
        Returns
        -------
        meritData : If successful, it is the number of operands in the merit
                    function
                    If meritData = -1, saving failed.

        See also: zOptimize, zLoadMerit
        """
        isAbsPath = _os.path.isabs(fileName)
        isRightExt = _os.path.splitext(fileName)[1] in ('.mf','.MF')
        if isAbsPath and isRightExt:
            cmd = "SaveMerit,{}".format(fileName)
            reply = self._sendDDEcommand(cmd)
            return int(float(reply.rstrip()))
        else:
            return -1

    def zSaveTolerance(self, fileName):
        """Saves the tolerances of the current lens to a file.

        saveTolerance(fileName)->numTolOperands

        Parameters
        ----------
        fileName : (string) file name of the file to save the tolerance data.
                   If no path is provided in the filename, the <data>\Tolerance
                   folder is assumed. Although it is not enforced, it is
                   useful to use ".tol" as extension.

        Returns
        -------
        numTolOperands : (integer) number of tolerance operands saved.

        See also zLoadTolerance.
        """
        cmd = "SaveTolerance,{}".format(fileName)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply.rstrip()))

    def zSetAperture(self,surfNum,aType,aMin,aMax,xDecenter=0,yDecenter=0,
                                                            apertureFile =' '):
        """Sets aperture details at a ZEMAX lens surface (surface data dialog box).

        zSetAperture(surfNum,aType,aMin,aMax,[xDecenter,yDecenter,apertureFile])
                                           -> apertureInfo

        Parameters
        ----------
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
                       see "Aperture type and other aperture controls" for more
                       details.
        xDecenter    : amount of decenter from current optical axis (lens units)
        yDecenter    : amount of decenter from current optical axis (lens units)
        apertureFile : a text file with .UDA extention. see "User defined
                       apertures and obscurations" in ZEMAX manual for more details.

        Returns
        -------
        apertureInfo: apertureInfo is a tuple containing the following:
            aType     : (see above)
            aMin      : (see above)
            aMax      : (see above)
            xDecenter : (see above)
            yDecenter : (see above)

        Example:
        -------
        apertureInfo = zSetAperture(2,1,5,10,0.5,0,'apertureFile.uda')
        or
        apertureInfo = zSetAperture(2,1,5,10)

        See also zGetAperture()
        """
        cmd  = ("SetAperture,{sN:d},{aT:d},{aMn:1.20g},{aMx:1.20g},{xD:1.20g},"
                "{yD:1.20g},{aF}".format(sN=surfNum,aT=aType,aMn=aMin,aMx=aMax,
                 xD=xDecenter,yD=yDecenter,aF=apertureFile))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        apertureInfo = tuple([float(elem) for elem in rs])
        return apertureInfo

    def zSetBuffer(self, bufferNumber, textData):
        """Used to store client specific data with the window being created or
        updated. The buffer data can be used to store user selected options,
        instead of using the settings data on the command line of the zMakeTextWindow
        or zMakeGraphicWindow functions. The data must be in a string format.

        zSetBuffer(bufferNumber, textData)->status

        Parameters
        ----------
        bufferNumber  : (integer) numbers between 0 and 15 inclusive (for 16
                       buffers provided)
        textData      : (string) is the only text that is stored, maximum of
                      240 characters

        Returns
        -------
        status        : 0 if successful, else -1

        Note
        -----
        The buffer data is not associated with any particular window until either
        the  zMakeTextWindow() or zMakeGraphicWindow() functions are issued. Once
        ZEMAX receives the MakeTextWindow or MakeGraphicWindow item, the buffer
        data is then copied to the appropriate window memory, and then may later
        be retrieved from that window's buffer using zGetBuffer() function.

        See also zGetBuffer.
        """
        if (0 < len(textData) < 240) and (0 <= bufferNumber < 16):
            cmd = "SetBuffer,{:d},{}".format(bufferNumber,str(textData))
            reply = self._sendDDEcommand(cmd)
            return 0 if 'OK' in reply.rsplit() else -1
        else:
            return -1

    def zSetConfig(self,configNumber):
        """Switches the current configuration number (selected column in the MCE),
        and updates the system.

        zSetConfig(configNumber)->(currentConfig, numberOfConfigs, error)

        Parameters
        ----------
        configNumber : The configuration (column) number to set current

        Returns
        -------
        3-tuple containing the following elements:
         currentConfig    : current configuration (column) number in MCE
                            1 <= currentConfig <= numberOfConfigs
         numberOfConfigs  : number of configs (columns).
         error            : 0  = successful; new current config is traceable
                            -1 = failure

        See also zGetConfig. Use zInsertConfig to insert new configuration in the
        multi-configuration editor.
        """
        reply = self._sendDDEcommand("SetConfig,{:d}".format(configNumber))
        rs = reply.split(',')
        return tuple([int(elem) for elem in rs])

    def zSetExtra(self,surfaceNumber,columnNumber,value):
        """Sets extra surface data (value) in the Extra Data Editor for the surface
        indicatd by surfaceNumber.

        zSetExtra(surfaceNumber,columnNumber,value)->retValue

        Parameters
        ----------
        surfaceNumber : (integer) surface number
        columnNumber  : (integer) column number
        value         : (float) value

        Returns
        -------
        retValue      : (float) numeric data value

        See also zGetExtra
        """
        cmd = ("SetExtra,{:d},{:d},{:1.20g}"
               .format(surfaceNumber,columnNumber,value))
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zSetField(self,n,arg1,arg2,arg3=1.0,vdx=0.0,vdy=0.0,vcx=0.0,vcy=0.0,van=0.0):
        """Sets the field data for a particular field point.

        There are 2 ways of using this function:

            zSetField(0, fieldType,totalNumFields,fieldNormalization)-> fieldData
             OR
            zSetField(n,xf,yf [,wgt,vdx,vdy,vcx,vcy,van])-> fieldData

        Parameters
        ----------
        [if n == 0]:
            0         : to set general field parameters
            arg1 : the field type
                  0 = angle, 1 = object height, 2 = paraxial image height, and
                  3 = real image height
            arg2 : total number of fields
            arg3 : normalization type [0=radial, 1=rectangular(default)]

        [if 0 < n <= number of fields]:
            arg1 (fx),arg2 (fy) : the field x and field y values
            arg3 (wgt)          : field weight (default = 1.0)
            vdx,vdy,vcx,vcy,van : vignetting factors (default = 0.0), See below.

        Returns
        -------
        [if n=0]: fieldData is a tuple containing the following
            type                 : integer (0=angles in degrees, 1=object height
                                            2=paraxial image height,
                                            3=real image height)
            number               : number of fields currently defined
            max_x_field          : values used to normalize x field coordinate
            max_y_field          : values used to normalize y field coordinate
            normalization_method : field normalization method (0=radial, 1=rectangular)

        [if 0 < n <= number of fields]: fieldData is a tuple containing the following
            xf     : the field x value
            yf     : the field y value
            wgt    : field weight
            vdx    : decenter x vignetting factor
            vdy    : decenter y vignetting factor
            vcx    : compression x vignetting factor
            vcy    : compression y vignetting factor
            van    : angle vignetting factor

        Note
        -----
        The returned tuple's content and structure is exactly same as that
        of zGetField()

        See also zGetField()
        """
        if n:
            cmd = ("SetField,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                   ",{:1.20g},{:1.20g},{:1.20g}"
                   .format(n,arg1,arg2,arg3,vdx,vdy,vcx,vcy,van))
        else:
            cmd = ("SetField,{:d},{:d},{:d},{:.0f}".format(0,arg1,arg2,arg3))

        reply = self._sendDDEcommand(cmd)
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

        Parameters
        ----------
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

        Returns
        -------
        oFieldDataTuple: the output field data tuple is also a N-D tuple similar
                         to the iFieldDataTuple, except that for each field location
                         all 8 field parameters are returned.

        See also zSetField(), zGetField(), zGetFieldTuple()
        """
        fieldCount = len(iFieldDataTuple)
        if not 0 < fieldCount <= 12:
            raise ValueError('Invalid number of fields')
        cmd = ("SetField,{:d},{:d},{:d},{:d}"
              .format(0,fieldType,fieldCount,fNormalization))
        self._sendDDEcommand(cmd)
        oFieldDataTuple = [ ]
        for i in range(fieldCount):
            fieldData = self.zSetField(i+1,*iFieldDataTuple[i])
            oFieldDataTuple.append(fieldData)
        return tuple(oFieldDataTuple)

    def zSetFloat(self):
        """Sets all surfaces without surface apertures to have floating apertures.
        Floating apertures will vignette rays which trace beyond the semi-diameter.

        zSetFloat()->status

        Parameters
        ----------
        None

        Returns
        -------
        status : 0 = success, -1 = fail
        """
        retVal = -1
        reply = self._sendDDEcommand('SetFloat')
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zSetLabel(self,surfaceNumber,label):
        """This command associates an integer label with the specified surface.
        The label will be retained by ZEMAX as surfaces are inserted or deleted
        around the target surface.

        zSetLabel(surfaceNumber, label)->assignedLabel

        Parameters
        ----------
        surfaceNumber : (integer) the surface number
        label         : (integer) the integer label

        Returns
        -------
        assignedLabel : (integer) should be equal to label

        See also zGetLabel, zFindLabel
        """
        reply = self._sendDDEcommand("SetLabel,{:d},{:d}"
                                          .format(surfaceNumber,label))
        return int(float(reply.rstrip()))

    def zSetMacroPath(self,macroFolderPath):
        """Set the full path name to the macro folder

        zSetMacroPath(macroFolderPath)->status

        Parameters
        ----------
        macroFolderPath : (string) full-path name of the macro folder path. Also,
                        this folder path should match the folder path specified
                        for Macros in the Zemax Preferences setting.

        Returns
        -------
        status : 0 = success, -1 = failure

        Note
        ----
        Use this method to set the full-path name of the macro folder
        path if it is different from the default path at <data>/Macros

        See also zExecuteZPLMacro
        """
        if _os.path.isabs(macroFolderPath):
            self.macroPath = macroFolderPath
            return 0
        else:
            return -1

    def zSetMulticon(self, config, *multicon_args):
        """Set data or operand type in the multi-configuration editior. Note that
        there are 2 ways of using this function.

        USAGE TYPE - I
        ==============
        If `config` is non-zero, then the function is used to set data in the
        MCE using the following syntax:

        `zSetMulticon(config,row,value,status,pickuprow,
                      pickupconfig,scale,offset)->multiConData`

        Example: `multiConData = zSetMulticon(1,5,5.6,0,0,0,1.0,0.0)`

        Parameters
        ----------
        config        : (int) configuration number (column)
        row           : (int) row or operand number
        value         : (float) value to set
        status        : (int) see below
        pickuprow     : (int) see below
        pickupconfig  : (int) see below
        scale         : (float)
        offset        : (float)

        Returns
        -------
        multiConData is a 8-tuple whose elements are:
        (value,num_config,num_row,status,pickuprow,pickupconfig,scale,offset)

        The `status` integer is 0 for fixed, 1 for variable, 2 for pickup, and 3
        for thermal pickup. If `status` is 2 or 3, the `pickuprow` and `pickupconfig`
        values indicate the source data for the pickup solve.


        USAGE TYPE - II
        ===============
        If the `config` = 0 , `zSetMulticon` may be used to set the operand type
        and number data using the following syntax:

        `zSetMulticon(0,row,operand_type,number1,number2,number3)-> multiConData`

        Example: `multiConData = zSetMulticon(0,5,'THIC',15,0,0)`

        Parameters
        ----------
        config       : 0
        row          : (int) row or operand number in the MCE
        operand_type : (string) operand type, such as 'THIC', 'WLWT', etc.
        number1      : (int)
        number2      : (int)
        number3      : (int)
                       Please refer to "SUMMARY OF MULTI-CONFIGURATION OPERANDS"
                       in the Zemax manual.

        Returns
        -------
        multiConData is a 4-tuple whose elements are:
        (operand_type,number1,number2,number3)

        NOTE:
        -----
        1. If there are current operands in the MCE, it is recommended to first
           use `zInsertMCO` to insert a row and then use `zSetMulticon(0,...)`. For
           example use `zInsertMCO(5)` and then use `zSetMulticon(0,5,'THIC',15,0,0)`.
           If a row is not inserted first, then existing rows may be overwritten.
        2. The functions raises an exception if it determines the arguments
           to be invalid.

        See also `zGetMulticon`
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
        reply = self._sendDDEcommand(cmd)
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

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        code : integer
            integer code
        data : string/integer/float
            data to set NSC object

            Refer table nsc-object-data-codes_ in the docstring of
            ``zGetNSCObjectData()`` for ``code`` and ``data`` specific details.

        Returns
        -------
        nscObjectData : string/integer/float
            the returned data (same as returned by ``zGetNSCObjectData()``)
            depends on the ``code``. If the command fails, it returns ``-1``.
            Refer table nsc-object-data-codes_.

        See Also
        --------
        zGetNSCObjectData(), zSetNSCObjectFaceData()
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
        reply = self._sendDDEcommand(cmd)
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

    def zSetNSCObjectFaceData(self, surfNumber, objNumber, faceNumber, code, data):
        """Sets the various data for NSC object faces

        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        faceNumber : integer
            face number
        code : integer
            integer code
        data : float/integer/string
            data to set NSC object face

            Refer table nsc-object-face-data-codes_ in the docstring of
            ``zGetNSCObjectData()`` for ``code`` and ``data`` specific details.

        Returns
        -------
        nscObjFaceData  : string/integer/float
            the returned data (same as returned by ``zGetNSCObjectFaceData()``)
            depends on the ``code``. If the command fails, it returns ``-1``.
            Refer table nsc-object-face-data-codes_.

        See Also
        --------
        zGetNSCObjectFaceData()
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
        reply = self._sendDDEcommand(cmd)
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

        Parameters
        ----------
        surfNumber      : (integer) surface number. Use 1 if
                          the program mode is Non-Sequential.
        objNumber       : (integer) object number
        parameterNumber : (integer) parameter number
        data            : (float) new numeric value for the parameterNumber

        Returns
        -------
        nscParaVal     : (float) parameter value

        See also zGetNSCParameter
        """
        cmd = ("SetNSCParameter,{:d},{:d},{:d},{:1.20g}"
              .format(surfNumber,objNumber,parameterNumber,data))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zSetNSCPosition(self,surfNumber,objectNumber,code,data):
        """Returns the position data for NSC objects.

        zSetNSCPosition(surfNumber,objectNumber,code,data)->nscPosData

        Parameters
        ----------
        surfNumber   : (integer) surface number. Use 1 if
                       the program mode is Non-Sequential.
        objectNumber : (integer) object number
        code         : (integer) 1-7 for x, y, z, tilt-x, tilt-y, tilt-z, and
                       material, respectively.
        data         : numeric (float) for codes 1-6, string for material (code-7)

        Returns
        -------
        nscPosData is a 7-tuple containing x,y,z,tilt-x,tilt-y,tilt-z,material

        See also zGetNSCPosition
        """
        if code == 7:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{}"
            .format(surfNumber,objectNumber,code,data))
        else:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{:1.20g}"
            .format(surfNumber,objectNumber,code,data))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPosData = -1
        else:
            nscPosData = tuple([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                    for i in range(len(rs))])
        return nscPosData

    def zSetNSCProperty(self, surfaceNumber, objectNumber, faceNumber, code, value):
        """Sets a numeric or string value to the property pages of objects
        defined in the NSC editor. It mimics the ZPL function NPRO.


        Parameters
        ----------
        surfaceNumber : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objectNumber : integer
            the NSC ojbect number
        faceNumber : integer
            face number
        code : integer
            for the specific code
        value : string/integer/float
            value to set NSC property

            Refer table nsc-property-codes_ in the docstring of ``zGetNSCProperty()``
            for ``code`` and ``value`` specific details.

        Returns
        -------
        nscPropData : string/float/integer
            the returned data (same as returned by ``zGetNSCProperty()``) depends
            on the ``code``. If the command fails, it returns ``-1``.
            Refer table nsc-property-codes_.

        See Also
        --------
        zGetNSCProperty()
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
        reply = self._sendDDEcommand(cmd)
        nscPropData = _process_get_set_NSCProperty(code,reply)
        return nscPropData

    def zSetNSCSettings(self,nscSettingsData):
        """Sets the maximum number of intersections, segments, nesting level,
        minimum absolute intensity, minimum relative intensity, glue distance,
        miss ray distance, and ignore errors flag used for NSC ray tracing.

        zSetNSCSettings(nscSettingsData)->nscSettingsDataRet

        Parameters
        ---------
        nscSettingsData is an 8-tuple with the following elements
        maxInt     : (integer) maximum number of intersections
        maxSeg     : (integer) maximum number of segments
        maxNest    : (integer) maximum nesting level
        minAbsI    : (float) minimum absolute intensity
        minRelI    : (float) minimum relative intensity
        glueDist   : (float) glue distance
        missRayLen : (float) miss ray distance
        ignoreErr  : (integer) 1 if true, 0 if false

        Returns
        -------
        nscSettingsDataRet is also an 8-tuple with the same elements as
        nscSettingsData.

        NOTE:
        -----
        Since the `maxSeg` value may require large amounts of RAM, verify
        that the new value was accepted by checking the returned tuple.

        See also zGetNSCSettings
        """
        (maxInt,maxSeg,maxNest,minAbsI,minRelI,glueDist,missRayLen,
                                                 ignoreErr) = nscSettingsData
        cmd = ("SetNSCSettings,{:d},{:d},{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:d}"
        .format(maxInt,maxSeg,maxNest,minAbsI,minRelI,glueDist,missRayLen,ignoreErr))
        reply = str(self._sendDDEcommand(cmd))
        rs = reply.rsplit(",")
        nscSettingsData = [float(rs[i]) if i in (3,4,5,6) else int(float(rs[i]))
                                                        for i in range(len(rs))]
        return tuple(nscSettingsData)

    def zSetNSCSolve(self, surfaceNumber, objectNumber, parameter, solveType,
                     pickupObject=0, pickupColumn=0, scale=0, offset=0):
        """Sets the solve type on NSC position and parameter data.

        zSetNSCSolve(surfaceNumber, objectNumber, parametersolveType,
                     pickupObject, pickupColumn, scale, offset) -> nscSolveData

        Parameters
        ----------
        surfaceNumber  : (integer) surface number. Use 1 if the program mode is
                         Non-Sequential.
        objectNumber   : (integer) object number
        parameter      : -1 = data for x data
                         -2 = data for y data
                         -3 = data for z data
                         -4 = data for tilt x data
                         -5 = data for tilt y data
                         -6 = data for tilt z data
                          n > 0  = data for the nth parameter
        solveType      : 0 = fixed, 1 = variable, 2 = pickup
        pickupObject   : if solveType = 2, pickup object number
        pickupColumn   : if solveType = 2, pickup column number (0 for current column)
        scale          : if solveType = 2, scale factor
        offset         : if solveType = 2, offset

        Returns
        -------
        nscSolveData : 5-tuple containing
                         (status, pickupObject, pickupColumn, scaleFactor, offset)
                         The status value is 0 for fixed, 1 for variable, and 2
                         for a pickup solve.
                         Only when the staus is a pickup solve is the other data
                         meaningful.
                       -1 if it a BAD COMMAND

        See also: zGetNSCSolve
        """
        nscSolveData = -1
        args1 = "{:d},{:d},{:d},".format(surfaceNumber, objectNumber, parameter)
        args2 = "{:d},{:d},{:d},".format(solveType, pickupObject, pickupColumn)
        args3 = "{:1.20g},{:1.20g}".format(scale, offset)
        cmd = ''.join(["SetNSCSolve,",args1,args2,args3])
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if 'BAD COMMAND' not in rs:
            nscSolveData = tuple([float(e) if i in (3,4) else int(float(e))
                                 for i,e in enumerate(rs.split(","))])
        return nscSolveData

    def zSetPrimaryWave(self,primaryWaveNumber):
        """Sets the wavelength data in the ZEMAX DDE server. This function emulates
        the function "zSetPrimaryWave()" of the MZDDE toolbox.

        zSetPrimaryWave(primaryWaveNumber) -> waveData

        Parameters
        ----------
        primaryWaveNumber: the wave number to set as primary

        Returns
        -------
        waveData is a tuple containing the following:
          primary : number indicating the primary wavelength (integer)
          number  : number of wavelengths currently defined (integer).

        Note
        ----
        The returned tuple is exactly same in structure and contents to that
        returned by zGetWave(0).

        See also zSetWave(), zSetWave(), zSetWaveTuple(), zGetWaveTuple().
        """
        waveData = self.zGetWave(0)
        cmd = "SetWave,{:d},{:d},{:d}".format(0,primaryWaveNumber,waveData[1])
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        waveData = tuple([int(elem) for elem in rs])
        return waveData

    def zSetOperand(self, row, column, value):
        """Sets the operand data in the Merit Function Editor

        Parameters
        ----------
        row : integer
            row operand number in the MFE
        column : integer
            column number
        value : string/integer/float
            the type of ``value`` depends on the ``column`` number

            Refer to the column-operand-data_ table (in the docstring of
            ``zGetOperand()`` for the column-value mapping)

        Returns
        -------
        operandData : string/integer/float
            the value set in the MFE cell. Refer table column-operand-data_.

        Notes
        -----
        1. To update the merit function after calling ``zSetOperand()``,
           call ``zOptimize()`` with the number of cycles set to -1.
        2. Use ``zInsertMFO()`` to insert additional rows, before calling
           ``zSetOperand()``.

        See Also
        --------
        zGetOperand(), zOptimize(), zInsertMFO()
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
        reply = self._sendDDEcommand(cmd)
        return _process_get_set_Operand(column, reply)

    def zSetPolState(self,nlsPolarized,Ex,Ey,Phx,Phy):
        """Sets the default polarization state. These parameters correspond to
        the Polarization tab under the General settings.

        zSetPolState()->polStateData

        Parameters
        ----------
        nlsPolarized : (integer) if nlsPolarized > 0, then default polarization
                       state is unpolarized.
        Ex           : (float) normalized electric field magnitude in x direction
        Ey           : (float) normalized electric field magnitude in y direction
        Phax         : (float) relative phase in x direction in degrees
        Phay         : (float) relative phase in y direction in degrees

        Returns
        -------
        polStateData is 5-tuple containing nlsPolarized, Ex, Ey, Phax, and Phay

        Note
        ----
        The quantity Ex*Ex + Ey*Ey should have a value of 1.0 although any
        values are accepted.

        See also zGetPolState.
        """
        cmd = ("SetPolState,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                .format(nlsPolarized,Ex,Ey,Phx,Phy))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rsplit(",")
        polStateData = [int(float(elem)) if i==0 else float(elem)
                                       for i,elem in enumerate(rs[:-1])]
        return tuple(polStateData)

    def zSetSettingsData(self, number, data):
        """This function sets the settings data used by a window in temporary
        storage before calling zMakeGraphicWindow or zMakeTextWindow. The data
        may be retrieved using zGetSettingsData.

        zSetSettingsData(number, data)->settingsData

        Parameters
        ----------
        number  : (integer) Currently, only number = 0 is supported. This number
                  may be used to expand the feature in the future.
        data    :

        Returns
        -------
        settingsData: (string)

        Note
        ----
        Please refer to "How ZEMAX calls the client" in the Zemax manual.
        See also zGetSettingsData.
        """
        cmd = "SettingsData,{:d},{}".format(number,data)
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())

    def zSetSolve(self, surfaceNumber, code, *solveData):
        """Sets data for solves and/or pickups on the surface with number
        `surfaceNumber`.

        `zSetSolve(surfaceNumber, code, *solveData)->solveData`

        also

        `zSetSolve(surfaceNumber, code, solvetype [, arg1, arg2, arg3, arg4])->solveData`

        Parameters
        ----------
        surfaceNumber : (integer) surface number for which the solve is to be set.
        code          : (integer) indicating which surface parameter the solve
                         data is for, such as curvature, thickness, glass, conic,
                         semi-diameter, etc. (see the table below)
        *solveData    : There are two ways of passing this parameter.
                        1. As a sequence of arguments:
                            solvetype, param1, param2, param3,and param4
                        2. As a tuple/list of the above arguments preceded by the
                           `*` operator to flatten/splatter the tuple/list (see example below).

                        The exact nature of the parameters depend on the `code` value
                        according to the following table.

        ------------------------------------------------------------------------
        code                    -  Solve data (solveData) & Returned data format
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

        Returns
        -------
        solveData     : a tuple, depending on the code value according to the
                        above table (same return as `zGetSolve`), if successful,
                        -1 if the command is a 'BAD COMMAND'

        Note
        ----
        The `solvetype` is an integer code, & the parameters have meanings
        that depend upon the solve type; see the chapter "SOLVES" in the Zemax
        manual for details.

        Example
        -------
        To set a solve on the curvature (0) of surface number 6 such that the
        Marginal Ray angle (2) value is 0.1, the following are equivalent:

        `sdata = ln.zSetSolve(6, 0, *(2, 0.1))`
        `sdata = ln.zSetSolve(6, 0,   2, 0.1 )`

        See also `zGetSolve`, `zGetNSCSolve`, `zSetNSCSolve`.
        """
        if not solveData:
            print("Error [zSetSolve] No solve data passed.")
            return -1
        try:
            if code == 0:          # Solve specified on CURVATURE
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
        reply = self._sendDDEcommand(cmd)
        solveData = _process_get_set_Solve(reply)
        return solveData

    def zSetSurfaceData(self, surfaceNumber, code, value, arg2=None):
        """Sets surface data on a sequential lens surface.

        zSetSurfaceData(surfaceNum,code,value [, arg2])-> surfaceDatum

        Parameters
        ----------
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
        Code      - Datum to be set by `zSetSurfaceData`
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

        See also `zGetSurfaceData` and `ZemaxSurfTypes`
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
        reply = self._sendDDEcommand(cmd)
        if code in (0,1,4,7,9):
            surfaceDatum = reply.rstrip()
        else:
            surfaceDatum = float(reply)
        return surfaceDatum

    def zSetSurfaceParameter(self,surfaceNumber,parameter,value):
        """Set surface parameter data.

        zSetSurfaceParameter(surfaceNumber, parameter, value)-> parameterData

        Parameters
        ----------
        surfaceNumber  : (integer) surface number of the surface
        parameter      : (integer) parameter (Par in LDE) number being set
        value          : (float) value to set for the `parameter`

        Returns
        -------
        parameterData  : (float) the parameter value

        See also `zSetSurfaceData`, `zGetSurfaceParameter`
        """
        cmd = ("SetSurfaceParameter,{:d},{:d},{:1.20g}"
               .format(surfaceNumber,parameter,value))
        reply = self._sendDDEcommand(cmd)
        return float(reply)


    def zSetSystem(self, unitCode, stopSurf, rayAimingType, useEnvData,
                                              temp, pressure, globalRefSurf):
        """Sets a number of general systems property (General Lens Data)

        zSetSystem(unitCode,stopSurf,rayAimingType,useenvdata,
                     temp,pressure,globalRefSurf) -> systemData

        Parameters
        ----------
        unitCode      : lens units code (0,1,2,or 3 for mm, cm, in, or M)
        stopSurf      : the stop surface number
        rayAimingType : ray aiming type (0,1, or 2 for off, paraxial or real)
        useEnvData    : use environment data flag (0 or 1 for no or yes) [ignored]
        temp          : the current temperature
        pressure      : the current pressure
        globalRefSurf : the global coordinate reference surface number

        Returns
        -------
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

        Note
        -----
        The returned data structure is exactly similar to the data structure
        returned by the `zGetSystem` method.

        If you are interested in setting the system apeture, such as aperture type,
        aperture value, etc, use `zSetSystemAper`.

        See also `zGetSystem`, `zGetSystemAper`, `zSetSystemAper`, `zGetAperture`,
        `zSetAperture`
        """
        cmd = ("SetSystem,{:d},{:d},{:d},{:d},{:1.20g},{:1.20g},{:d}"
              .format(unitCode,stopSurf,rayAimingType,useEnvData,temp,pressure,
               globalRefSurf))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        systemData = tuple([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zSetSystemAper(self, aType, stopSurf, apertureValue):
        """Sets the lens system aperture and corresponding data.

        zSetSystemAper(aType, stopSurf, apertureValue)-> systemAperData

        Parameters
        ---------
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

        Returns
        -------
        systemAperData: systemAperData is a tuple containing the following
            aType              : (see above)
            stopSurf           : (see above)
            value              : (see above)

        Note: the returned tuple is the same as the returned tuple of `zGetSystemAper`

        See also, `zGetSystem`, `zGetSystemAper`
        """
        cmd = ("SetSystemAper,{:d},{:d},{:1.20g}"
               .format(aType,stopSurf,apertureValue))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        systemAperData = tuple([float(elem) if i==2 else int(float(elem))
                                for i, elem in enumerate(rs)])
        return systemAperData

    def zSetSystemProperty(self, code, value1, value2=0):
        """Sets system properties of the system.

        System properties such as system aperture, field, wavelength, and other data,
        based on the integer `code` passed.

        zSetSystemProperty(code)-> sysPropData

        Parameters
        ----------
        code        : (integer) value that defines the specific system property
                      to be set (see below).
        value1      : (integer/float/string) depending on `code`
        value2      : (integer/float), ignored if not used

        Returns
        -------
        sysPropData : Returned system property data. Either a string or numeric
                      data.

        This function mimics the ZPL function SYPR.
        ------------------------------------------------------------------------
        Code    Property (the values in the bracket are the expected returns)
        ------------------------------------------------------------------------
          4   - Adjust Index Data To Environment. (0:off, 1:on.)
         10   - Aperture Type code. (0:EPD, 1:IF/#, 2:ONA, 3:FBS, 4:PWF/#, 5:OCA)
         11   - Aperture Value. (stop surface semi-diameter if aperture type is
                FBS, else system aperture)
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

        NOTE: Currently Zemax returns just "0" for the codes: 102,103, 104,105,
              106,107,108,109, and 110. This is unexpected! So, PyZDDE will return
              the reply (string) as is for the user to handle. The zSetSystemProperty
              functions as expected nevertheless.

        See also `zGetSystemProperty`.
        """
        cmd = "SetSystemProperty,{c:d},{v1},{v2}".format(c=code,v1=value1,v2=value2)
        reply = self._sendDDEcommand(cmd)
        sysPropData = _process_get_set_SystemProperty(code,reply)
        return sysPropData

    def zSetTol(self,operandNumber,col,value):
        """Sets the tolerance operand data.

        zSetTol(operandNumber,col,value)-> toleranceData

        Parameters
        ----------
        operandNumber : (integer) tolerance operand number (row number in the
                        tolerance editor, when greater than 0)
        col           : (integer) 1 for tolerance Type
                        (integer) 2-4 for int1-int3
                        (integer) 5 for min
                        (integer) 6 for max
        value         : 4-character string (tolerancing operand code) if col==1,
                        else float value to set

        Returns
        -------
        toleranceData. It is a number or a 6-tuple, depending upon `operandNumber`
        as follows:

        * if operandNumber == 0, toleranceData = number where `number` is the number of
          tolerance operands defined.

        * if operandNumber > 0, toleranceData = (tolType, int1, int2, min, max, int3)

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
        reply = self._sendDDEcommand(cmd)
        if operandNumber == 0: # returns just the number of operands
            return int(float(reply.rstrip()))
        else:
            return _process_get_set_Tol(operandNumber,reply)
        # FIX !!! currently, I am not able to set more than 1 row in the tolerance
        # editor, through this command. I don't find anything like zInsertTol ...
        # A similar function exist for Multi-Configuration editor (zInsertMCO) and
        # for Multi-function editor (zInsertMFO). May need to contact Zemax Support.

    def zSetTolRow(self,operandNumber,tolType,int1,int2,int3,minT,maxT):
        """Helper function to set all the elements of a row (given by operandNumber)
        in the tolerance editor.

        zSetTolRow(operandNumber,tolType,int1,int2,int3,minT,maxT)->tolData

        Parameters
        ----------
        operandNumber  : (integer greater than 0) tolerance operand number (row
                         number in the tolerance editor)
        tolType        : 4-character string (tolerancing operand code)
        int1           : (integer) int1 parameter
        int2           : (integer) int2 parameter
        int3           : (integer) int3 parameter
        minT           : (float) minimum value
        maxT           : (float) maximum value

        Returns
        -------
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

        Parameters
        ----------
        bufferCode : (integer) The buffercode is an integer value provided
                     by ZEMAX to the client that uniquely identifies the
                     correct lens.
        dataNumber : (integer)
        data       : (float) data item number being passed

        Returns
        -------
          ?

        Note
        -----
        After the last data item has been sent, the buffer must be closed
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
        reply = self._sendDDEcommand(cmd)
        return _regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zSetVig(self):
        """Request Zemax to set the vignetting factors automatically. Calling this
        function is equivalent to clicking the "Set Vig" button from the "Field Data"
        window. For more information on how Zemax calculates the vignetting factors
        automatically, please refer to "Vignetting factors" under the "Systems Menu"
        chapter in the Zemax Manual.

        zSetVig()->retVal

        Parameters
        ----------
        None

        Returns
        -------
        retVal  : 0 = success, -1 = fail
        """
        retVal = -1
        reply = self._sendDDEcommand("SetVig")
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zSetWave(self,n,arg1,arg2):
        """Sets the wavelength data in the ZEMAX DDE server.

        There are 2 ways to use this function:
            zSetWave(0,primary,number) -> waveData
             OR
            zSetWave(n,wavelength,weight) -> waveData

        Parameters
        ----------
        [if n==0]:
            0             : if n=0, the function sets general wavelength data
            primary (arg1): primary wavelength value to set
            number (arg2) : total number of wavelengths to set

        [if 0 < n <= number of wavelengths]:
            n                : wavelength number to set
            wavelength (arg1): wavelength in micrometers (floating)
            weight    (arg2) : weight (floating)

        Returns
        -------
        if n==0: waveData is a tuple containing the following:
            primary : number indicating the primary wavelength (integer)
            number  : number of wavelengths currently defined (integer).
        elif 0 < n <= number of wavelengths: waveData consists of:
            wavelength : value of the specific wavelength (floating point)
            weight     : weight of the specific wavelength (floating point)

        Notes
        -----
        The returned tuple is exactly same in structure and contents to that
        returned by zGetWave().

        See Also
        --------
        zGetWave(), zSetPrimaryWave(), zSetWaveTuple(), zGetWaveTuple()
        """
        if n:
            cmd = "SetWave,{:d},{:1.20g},{:1.20g}".format(n,arg1,arg2)
        else:
            cmd = "SetWave,{:d},{:d},{:d}".format(0,arg1,arg2)

        reply = self._sendDDEcommand(cmd)
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

        Parameters
        ----------
        iWaveDataTuple: the input wave data tuple is a 2D tuple with the first
                        dimension (first subtuple) containing the wavelengths and
                        the second dimension containing the weights like so:
                        ((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))

        The first wavelength (wave1) is assigned to be the primary wavelength.
        To change the primary wavelength use zSetWavePrimary()

        Returns
        -------
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
            reply = self._sendDDEcommand(cmd)
            rs = reply.split(',')
            oWaveDataTuple[0].append(float(rs[0])) # store the wavelength
            oWaveDataTuple[1].append(float(rs[1])) # store the weight
        return (tuple(oWaveDataTuple[0]),tuple(oWaveDataTuple[1]))

    def zWindowMaximize(self,windowNumber=0):
        """Maximize the main ZEMAX window or any analysis window ZEMAX currently
        displayed.

        zWindowMaximize(windowNumber)->retVal

        Parameters
        ----------
        windowNumber  : (integer) window number. use 0 for the main ZEMAX window

        Returns
        -------
        retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowMaximize,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowMinimize(self,windowNumber=0):
        """Minimize the main ZEMAX window or any analysis window ZEMAX currently
        displayed.

        zWindowMinimize(windowNumber)->retVal

        Parameters
        -----------
        windowNumber  : (integer) window number. use 0 for the main ZEMAX window

        Returns
        -------
        retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowMinimize,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowRestore(self,windowNumber=0):
        """Restore the main ZEMAX window or any analysis window to it's previous
        size and position.

        zWindowRestore(windowNumber)->retVal

        Parameters
        ----------
        windowNumber  : (integer) window number. use 0 for the main ZEMAX window

        Returns
        -------
        retVal   : 0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowRestore,{:d}".format(windowNumber))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

# ****************************************************************
#                      EXTRA FUNCTIONS
# ****************************************************************
    def zSpiralSpot(self, hx, hy, waveNum, spirals, rays, mode=0):
        """returns positions and intensity of rays traced in a spiral over the
        entrance pupil to the image surface.

        zSpiralSpot(hx,hy,waveNum,spirals,rays[,mode])->(x,y,z,intensity)

        The final destination of the rays is the image surface. This function
        imitates its namesake from MZDDE toolbox. (Note: unlike the spiralSpot
        of MZDDE, there is no need to call zLoadLens() before calling zSpiralSpot()).

        Parameters
        ----------
        hx      : normalized field height along x axis
        hy      : normalized field height along y axis
        waveNum : wavelength number as in the wavelength data editor
        mode    : 0 = real, 1 = paraxial

        Returns
        -------
        rayInfo : 4-tuple = (x,y,z,intensity)
        """
        # Calculate the ray pattern on the pupil plane
        pi, cos, sin = _math.pi, _math.cos, _math.sin
        lastAng = spirals*2*pi
        delta_t = lastAng/(rays-1)
        theta = lambda dt, rays: (i*dt for i in range(rays))
        r = (i/lastAng for i in theta(delta_t, rays))
        pXY = ((r*cos(t), r*sin(t)) for r, t in _izip(r,theta(delta_t, rays)))
        x = [] # x-coordinate of the image surface
        y = [] # y-coordinate of the image surface
        z = [] # z-coordinate of the image surface
        intensity = [] # the relative transmitted intensity of the ray
        for px, py in pXY:
            rayTraceData = self.zGetTrace(waveNum, mode, -1, hx, hy, px, py)
            if rayTraceData[0] == 0:
                x.append(rayTraceData[2])
                y.append(rayTraceData[3])
                z.append(rayTraceData[4])
                intensity.append(rayTraceData[11])
            else:
                print("Raytrace Error")
                exit()
                # !!! FIX raise an error here
        return (x, y, z, intensity)

    def zLensScale(self,factor=2.0,ignoreSurfaces=None):
        """Scale the lens design by factor specified.

        zLensScale([factor,ignoreSurfaces])->ret

        Parameters
        ----------
        factor         : the scale factor. If no factor are passed, the design
                         will be scaled by a factor of 2.0
        ignoreSurfaces : (tuple) of surfaces that are not to be scaled. Such as
                         (0,2,3) to ignore surfaces 0 (object surface), 2 and
                         3. Or (OBJ,2, STO,IMG) to ignore object surface, surface
                         number 2, stop surface and image surface.
        Returns
        -------
        0 : success
        1 : success with warning
        -1: failure

        Notes
        -----
        1. WARNING: this function implementation is not yet complete.
            * Note all surfaces have been implemented
            * ignoreSurface option has not been implemented yet.

        Limitations:
        -----------
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
                    self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*pNum)*par)
                #Scale norm radius in the extra data editor
                epar2 = self.zGetExtra(surfNum,2) #Norm radius
                self.zSetExtra(surfNum,2,factor*epar2)
                #scale the coefficients of the Zernike Fringe polynomial terms in the EDE
                numBTerms = int(self.zGetExtra(surfNum,1))
                if numBTerms > 0:
                    for i in range(3,binSurMaxNum[surfName]): # scaling of terms 3 to 232, p^480
                                                              # for Binary1 and Binary 2 respectively
                        if i > numBTerms + 2: #(+2 because the terms starts from par 3)
                            break
                        else:
                            epar = self.zGetExtra(surfNum,i)
                            self.zSetExtra(surfNum,i,factor*epar)
            elif surfName == 'BINARY_3':
                #Scaling of parameters in the LDE
                par1 = self.zGetSurfaceParameter(surfNum,1) # R2
                self.zSetSurfaceParameter(surfNum,1,factor*par1)
                par4 = self.zGetSurfaceParameter(surfNum,4) # A2, need to scale A2 before A1,
                                                            # because A2>A1>0.0 always
                self.zSetSurfaceParameter(surfNum,4,factor*par4)
                par3 = self.zGetSurfaceParameter(surfNum,3) # A1
                self.zSetSurfaceParameter(surfNum,3,factor*par3)
                numBTerms = int(self.zGetExtra(surfNum,1))    #Max possible is 60
                for i in range(2,243,4):  #242
                    if i > 4*numBTerms + 1: #(+1 because the terms starts from par 2)
                        break
                    else:
                        par_r1 = self.zGetExtra(surfNum,i)
                        self.zSetExtra(surfNum,i,par_r1/factor**(i/2))
                        par_p1 = self.zGetExtra(surfNum,i+1)
                        self.zSetExtra(surfNum,i+1,factor*par_p1)
                        par_r2 = self.zGetExtra(surfNum,i+2)
                        self.zSetExtra(surfNum,i+2,par_r2/factor**(i/2))
                        par_p2 = self.zGetExtra(surfNum,i+3)
                        self.zSetExtra(surfNum,i+3,factor*par_p2)

            elif surfName == 'COORDBRK': #Coordinate break,
                par = self.zGetSurfaceParameter(surfNum,1) # decenter X
                self.zSetSurfaceParameter(surfNum,1,factor*par)
                par = self.zGetSurfaceParameter(surfNum,2) # decenter Y
                self.zSetSurfaceParameter(surfNum,2,factor*par)
            elif surfName == 'EVENASPH': #Even Asphere,
                for pNum in range(1,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    self.zSetSurfaceParameter(surfNum,pNum,
                                                         factor**(1-2.0*pNum)*par)
            elif surfName == 'GRINSUR1': #Gradient1
                par1 = self.zGetSurfaceParameter(surfNum,1) #Delta T
                self.zSetSurfaceParameter(surfNum,1,factor*par1)
                par3 = self.zGetSurfaceParameter(surfNum,3) #coeff of radial quadratic index
                self.zSetSurfaceParameter(surfNum,3,par3/(factor**2))
                par4 = self.zGetSurfaceParameter(surfNum,4) #index of radial linear index
                self.zSetSurfaceParameter(surfNum,4,par4/factor)
            elif surfName == 'GRINSUR9': #Gradient9
                par = self.zGetSurfaceParameter(surfNum,1) #Delta T
                self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'GRINSU11': #Grid Gradient surface with 1 parameter
                par = self.zGetSurfaceParameter(surfNum,1) #Delta T
                self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'PARAXIAL': #Paraxial
                par = self.zGetSurfaceParameter(surfNum,1) #Focal length
                self.zSetSurfaceParameter(surfNum,1,factor*par)
            elif surfName == 'PARAX_XY': #Paraxial XY
                par = self.zGetSurfaceParameter(surfNum,1) # X power
                self.zSetSurfaceParameter(surfNum,1,par/factor)
                par = self.zGetSurfaceParameter(surfNum,2) # Y power
                self.zSetSurfaceParameter(surfNum,2,par/factor)
            elif surfName == 'PERIODIC':
                par = self.zGetSurfaceParameter(surfNum,1) #Amplitude/ peak to valley height
                self.zSetSurfaceParameter(surfNum,1,factor*par)
                par = self.zGetSurfaceParameter(surfNum,2) #spatial frequency of oscillation in x
                self.zSetSurfaceParameter(surfNum,2,par/factor)
                par = self.zGetSurfaceParameter(surfNum,3) #spatial frequency of oscillation in y
                self.zSetSurfaceParameter(surfNum,3,par/factor)
            elif surfName == 'POLYNOMI':
                for pNum in range(1,5): # from Par 1 to Par 4 for x then Par 5 to Par 8 for y
                    parx = self.zGetSurfaceParameter(surfNum,pNum)
                    pary = self.zGetSurfaceParameter(surfNum,pNum+4)
                    self.zSetSurfaceParameter(surfNum,pNum,
                                                      factor**(1-2.0*pNum)*parx)
                    self.zSetSurfaceParameter(surfNum,pNum+4,
                                                         factor**(1-2.0*pNum)*pary)
            elif surfName == 'TILTSURF': #Tilted surface
                pass           #No parameters to scale
            elif surfName == 'TOROIDAL':
                par = self.zGetSurfaceParameter(surfNum,1) #Radius of rotation
                self.zSetSurfaceParameter(surfNum, 1, factor*par)
                for pNum in range(2,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    self.zSetSurfaceParameter(surfNum,pNum,
                                              factor**(1-2.0*(pNum-1))*par)
                #scale parameters from the extra data editor
                epar = self.zGetExtra(surfNum, 2)
                self.zSetExtra(surfNum, 2, factor*epar)
            elif surfName == 'FZERNSAG': # Zernike fringe sag
                for pNum in range(1,9): # from Par 1 to Par 8
                    par = self.zGetSurfaceParameter(surfNum,pNum)
                    self.zSetSurfaceParameter(surfNum,pNum,
                                              factor**(1-2.0*pNum)*par)
                par9 = self.zGetSurfaceParameter(surfNum,9) # decenter X
                self.zSetSurfaceParameter(surfNum,9,factor*par9)
                par10 = self.zGetSurfaceParameter(surfNum,10) # decenter Y
                self.zSetSurfaceParameter(surfNum,10,factor*par10)
                #Scale norm radius in the extra data editor
                epar2 = self.zGetExtra(surfNum,2) #Norm radius
                self.zSetExtra(surfNum,2,factor*epar2)
                #scale the coefficients of the Zernike Fringe polynomial terms in the EDE
                numZerTerms = int(self.zGetExtra(surfNum,1))
                if numZerTerms > 0:
                    epar3 = self.zGetExtra(surfNum,3) #Zernike Term 1
                    self.zSetExtra(surfNum,3,factor*epar3)
                    #Zernike terms 2,3,4,5 and 6 are not scaled.
                    for i in range(9,40): #scaling of Zernike terms 7 to 37
                        if i > numZerTerms + 2: #(+2 because the Zernike terms starts from par 3)
                            break
                        else:
                            epar = self.zGetExtra(surfNum,i)
                            self.zSetExtra(surfNum,i,factor*epar)
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


    def zCalculateHiatus(self, txtFile=None, keepFile=False):
        """Calculate the Hiatus.

        The hiatus, also known as the Null space, or nodal space, or the
        interstitium is the distance between the two principal planes.

        zCalculateHiatus([txtFile,keepFile])-> hiatus

        Parameters
        ----------
        txtFile : (optional, string) If passed, the prescription file
                          will be named such. Pass a specific txtFile if
                          you want to dump the file into a separate directory.
        keepFile        : (optional, bool) If false (default), the prescription
                          file will be deleted after use. If true, the file
                          will persist.
        Returns
        -------
        hiatus          : the value of the hiatus
        """
        if txtFile is not None:
            textFileName = txtFile
        else:
            cd = _os.path.dirname(_os.path.realpath(__file__))
            textFileName = _os.path.join(cd, "prescriptionFile.txt")
        ret = self.zGetTextFile(textFileName, 'Pre', "None", 0)
        assert ret == 0
        recSystemData = self.zGetSystem()
        numSurf = recSystemData[0]

        # The number of expected Principal planes in each Pre file is equal to the
        # number of wavelengths in the general settings of the lens design
        line_list = _readLinesFromFile(_openFile(textFileName))

        principalPlane_objSpace = 0.0
        principalPlane_imgSpace = 0.0
        hiatus = 0.0
        count = 0

        for line_num, line in enumerate(line_list):
            # Extract the image surface distance from the global ref sur (surface 1)
            sectionString = ("GLOBAL VERTEX COORDINATES, ORIENTATIONS,"
                             " AND ROTATION/OFFSET MATRICES:")
            if line.rstrip() == sectionString:
                ima_3 = line_list[line_num + numSurf*4 + 6]
                ima_z = float(ima_3.split()[3])

            # Extract the Principal plane distances.
            if "Principal Planes" in line and "Anti" not in line:
                principalPlane_objSpace += float(line.split()[3])
                principalPlane_imgSpace += float(line.split()[4])
                count +=1  #Increment (wavelength) counter for averaging

        # Calculate the average (for all wavelengths) of the principal plane distances
        if count > 0:
            principalPlane_objSpace = principalPlane_objSpace/count
            principalPlane_imgSpace = principalPlane_imgSpace/count
            #Calculate the hiatus (only if count > 0) as
            #hiatus = (img_surf_dist + img_surf_2_imgSpacePP_dist) - objSpacePP_dist
            hiatus = abs(ima_z + principalPlane_imgSpace - principalPlane_objSpace)

        if not keepFile:
            _deleteFile(textFileName)
        return hiatus

    def zGetPOP(self, settingsFile=None, displayData=False, txtFile=None, 
                keepFile=False, timeout=None):
        """returns Physical Optics Propagation (POP) data

        Parameters
        ----------
        settingsFile : string, optional
            * if passed, the POP will be called with this configuration file; 
            * if no ``settingsFile`` is passed, and config file ending with the
              same name as the lens file post fixed with "_pyzdde_POP.CFG" is 
              present, the settings from this file will be used;
            * if no ``settingsFile`` and no file name post-fixed with 
              "_pyzdde_POP.CFG" is found, but a config file with the same name 
              as the lens file is present, the settings from that file will be 
              used;
            * if no settings file is found, then a default settings will be used 
        displayData : bool
            if ``true`` the function returns the 2D display data; default 
            is ``false``
        txtFile : string, optional
            if passed, the POP data file will be named such. Pass a 
            specific ``txtFile`` if you want to dump the file into a 
            separate directory.
        keepFile : bool, optional 
            if false (default), the prescription file will be deleted after 
            use. If true, the file will persist.
        timeout : integer, optional
            timeout in seconds

        Returns
        -------
        popData : tuple
            popData is a 1-tuple continining just ``popInfo`` (see below) if 
            ``displayData`` is ``false`` (default). 
            If ``displayData`` is ``true``, ``popData`` is a 2-tuple containing 
            ``popInfo`` (a tuple) and ``powerGrid`` (a 2D list):

            popInfo : tuple
                peakIrradiance/ centerPhase : float
                    the peak irradiance is the maximum power per unit area 
                    at any point in the beam, measured in source units per 
                    lens unit squared. It returns center phase if the data
                    type is "Phase" in POP settings
                totalPower : float 
                    the total power, or the integral of the irradiance over 
                    the entire beam if data type is "Irradiance" in POP 
                    settings. This field is blank for "Phase" data
                fiberEfficiency_system : float
                    the efficiency of power transfer through the system
                fiberEfficiency_receiver : float
                    the efficiency of the receiving fiber
                coupling : float
                    the total coupling efficiency, the product of the system 
                    and receiver efficiencies
                pilotSize : float
                    the size of the gaussian beam at the surface
                pilotWaist : float
                    the waist of the gaussian beam
                pos : float
                    relative z position of the gaussian beam
                rayleigh : float 
                    the rayleigh range of the gaussian beam
                gridX : integer
                    the X-sampling
                gridY : interger
                    the Y-sampling
                widthX : float
                    width along X in lens units
                widthY : float
                    width along Y in lens units
        
            powerGrid : 2D list/ None
                a two-dimensional list of the powers in the analysis grid if
                ``displayData`` is ``true``

        Notes
        ----- 
        The function returns ``None`` for any field which was not found in POP 
        text file. This is most common in the case of ``fiberEfficiency_system``
        and ``fiberEfficiency_receiver`` as they need to be set explicitly in 
        the POP settings

        See Also
        -------- 
        zSetPOPSettings(), zModifyPOPSettings()   
        """
        cd = _os.path.dirname(_os.path.realpath(__file__))
        if txtFile != None:
            textFileName = txtFile
        else:
            textFileName = _os.path.join(cd, "popData.txt")
        # decide about what settings to use
        if settingsFile:
            cfgFile = settingsFile    
            getTextFlag = 1 
        else:
            f = _os.path.splitext(self.zGetFile())[0] + '_pyzdde_POP.CFG'
            if _checkFileExist(f): # use "*_pyzdde_POP.CFG" settings file 
                cfgFile = f
                getTextFlag = 1
            else: # use default settings file
                cfgFile = ''
                getTextFlag = 0
        
        ret = self.zGetTextFile(textFileName, 'Pop', cfgFile, getTextFlag)
        assert ret == 0
        # get line list
        line_list = _readLinesFromFile(_openFile(textFileName))
        
        # Get data type ... phase or Irradiance?
        find_irr_data = _getFirstLineOfInterest(line_list, 'POP Irradiance Data', 
                                                patAtStart=False) 
        data_is_irr = False if find_irr_data is None else True

        # Get the Grid size
        grid_line = line_list[_getFirstLineOfInterest(line_list, 'Grid size')]
        grid_x, grid_y = [int(i) for i in _re.findall(r'\d{2,5}', grid_line)]

        # Point spacing
        pts_line = line_list[_getFirstLineOfInterest(line_list, 'Point spacing')]
        pat = r'-?\d\.\d{4,6}[Ee][-\+]\d{3}'
        pts_x, pts_y =  [float(i) for i in _re.findall(pat, pts_line)]

        width_x = pts_x*grid_x
        width_y = pts_y*grid_y

        if data_is_irr:
            # Peak Irradiance and Total Power
            pat_i = r'-?\d\.\d{4,6}[Ee][-\+]\d{3}' # pattern for P. Irr, T. Pow,
            peakIrr, totPow = None, None
            pi_tp_line = line_list[_getFirstLineOfInterest(line_list, 'Peak Irradiance')]
            if pi_tp_line: # Transfer magnitude doesn't have Peak Irradiance info
                pi_info, tp_info = pi_tp_line.split(',')
                pi = _re.search(pat_i, pi_info)
                tp = _re.search(pat_i, tp_info)
                if pi:
                    peakIrr = float(pi.group())
                if tp:
                    totPow = float(tp.group())
        else:
            # Center Phase
            pat_p = r'-?\d+\.\d{4,6}' # pattern for Center Phase Info
            centerPhase = None
            cp_line = line_list[_getFirstLineOfInterest(line_list, 'Center Phase')]
            if cp_line: # Transfer magnitude / Phase doesn't have Center Phase info
                cp = _re.search(pat_p, cp_line)
                if cp:
                    centerPhase = float(cp.group())
        # Pilot_size, Pilot_Waist, Pos, Rayleigh [... available for both Phase and Irr data]
        pat_fe = r'\d\.\d{6}'   # pattern for fiber efficiency
        pat_pi = r'-?\d\.\d{4,6}[Ee][-\+]\d{3}' # pattern for Pilot size/waist
        pilotSize, pilotWaist, pos, rayleigh = None, None, None, None
        pilot_line = line_list[_getFirstLineOfInterest(line_list, 'Pilot')]
        p_size_info, p_waist_info, p_pos_info, p_rayleigh_info = pilot_line.split(',')
        p_size = _re.search(pat_pi, p_size_info)
        p_waist = _re.search(pat_pi, p_waist_info)
        p_pos = _re.search(pat_pi, p_pos_info)
        p_rayleigh = _re.search(pat_pi, p_rayleigh_info)
        if p_size:
            pilotSize = float(p_size.group())
        if p_waist:
            pilotWaist = float(p_waist.group())
        if p_pos:
            pos = float(p_pos.group())
        if p_rayleigh:
            rayleigh = float(p_rayleigh.group())

        # Fiber Efficiency, Coupling [... if enabled in settings]
        fibEffSys, fibEffRec, coupling = None, None, None
        effi_coup_line_num = _getFirstLineOfInterest(line_list, 'Fiber Efficiency')
        if effi_coup_line_num:
            efficiency_coupling_line = line_list[effi_coup_line_num]
            efs_info, fer_info, cou_info = efficiency_coupling_line.split(',')
            fes = _re.search(pat_fe, efs_info)
            fer = _re.search(pat_fe, fer_info)
            cou = _re.search(pat_fe, cou_info)
            if fes:
                fibEffSys = float(fes.group())
            if fer:
                fibEffRec = float(fer.group())
            if cou:
                coupling = float(cou.group())
        
        if displayData:
            # Get the 2D data
            pat = (r'(-?\d\.\d{4,6}[Ee][-\+]\d{3}\s*)' + r'{{{num}}}'
                   .format(num=grid_x))
            start_line = _getFirstLineOfInterest(line_list, pat)
            powerGrid = _get2DList(line_list, start_line, grid_y)
        
        if not keepFile:
            _deleteFile(textFileName)

        if data_is_irr: # Irradiance data
            popi = _co.namedtuple('POPinfo', ['peakIrr', 'totPow', 
                                              'fibEffSys', 'fibEffRec', 'coupling', 
                                              'pilotSize', 'pilotWaist', 'pos', 
                                              'rayleigh', 'gridX', 'gridY',
                                              'widthX', 'widthY' ])
            popInfo = popi(peakIrr, totPow, fibEffSys, fibEffRec, coupling, 
                           pilotSize, pilotWaist, pos, rayleigh, 
                           grid_x, grid_y, width_x, width_y)
        else: # Phase data
            popi = _co.namedtuple('POPinfo', ['cenPhase', 'blank', 
                                              'fibEffSys', 'fibEffRec', 'coupling', 
                                              'pilotSize', 'pilotWaist', 'pos', 
                                              'rayleigh', 'gridX', 'gridY',
                                              'widthX', 'widthY' ])
            popInfo = popi(centerPhase, None, fibEffSys, fibEffRec, coupling, 
                           pilotSize, pilotWaist, pos, rayleigh, 
                           grid_x, grid_y, width_x, width_y)
        if displayData:
            return (popInfo, powerGrid)
        else:
            return popInfo

    def zModifyPOPSettings(self, settingsFile, start_surf=None,  
                           end_surf=None, field=None, wave=None, auto=None, 
                           beamType=None, paramN=((),()), pIrr=None, tPow=None, 
                           sampx=None, sampy=None, srcFile=None, widex=None, 
                           widey=None, fibComp=None, fibFile=None, fibType=None, 
                           fparamN=((),()), ignPol=None, pos=None, tiltx=None, 
                           tilty=None):
        """modify an existing POP settings file

        Only those parameters with non-None or non-zero-length (in case of tuples)
        will be set.

        Parameters
        ---------- 
        settingsFile : string
            filename of the settings file including path and extension
        others :
            see the parameter definitions of ``zSetPOPSettings()``

        Returns
        ------- 
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each non-None
            parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid. 

        See Also
        -------- 
        zSetPOPSettings(), zGetPOP()
        """
        sTuple = []
        if (_os.path.isfile(settingsFile) and 
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if start_surf is not None:
            sTuple.append(self.zModifySettings(dst, "POP_START", start_surf))
        if end_surf is not None:
            sTuple.append(self.zModifySettings(dst, "POP_END", end_surf))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "POP_FIELD", field))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "POP_WAVE", wave))
        if auto is not None:
            sTuple.append(self.zModifySettings(dst, "POP_AUTO", auto))
        if beamType is not None:
            sTuple.append(self.zModifySettings(dst, "POP_BEAMTYPE", beamType))
        if paramN[0]:
            tst = []
            for i, j in _izip(paramN[0], paramN[1]):
                tst.append(self.zModifySettings(dst, "POP_PARAM{}".format(i), j))
            sTuple.append(tuple(tst))
        if pIrr is not None:
            sTuple.append(self.zModifySettings(dst, "POP_PEAKIRRAD", pIrr))
        if tPow is not None:
            sTuple.append(self.zModifySettings(dst, "POP_POWER", tPow))
        if sampx is not None:
            sTuple.append(self.zModifySettings(dst, "POP_SAMPX", sampx))
        if sampy is not None:
            sTuple.append(self.zModifySettings(dst, "POP_SAMPY", sampy))
        if srcFile is not None:
            sTuple.append(self.zModifySettings(dst, "POP_SOURCEFILE", srcFile))
        if widex is not None:
            sTuple.append(self.zModifySettings(dst, "POP_WIDEX", widex))
        if widey is not None:
            sTuple.append(self.zModifySettings(dst, "POP_WIDEY", widey))
        if fibComp is not None:
            sTuple.append(self.zModifySettings(dst, "POP_COMPUTE", fibComp))
        if fibFile is not None:
            sTuple.append(self.zModifySettings(dst, "POP_FIBERFILE", fibFile))
        if fibType is not None:
            sTuple.append(self.zModifySettings(dst, "POP_FIBERTYPE", fibType))
        if fparamN[0]:
            tst = []
            for i, j in _izip(fparamN[0], fparamN[1]):
                tst.append(self.zModifySettings(dst, "POP_FPARAM{}".format(i), j))
            sTuple.append(tuple(tst))
        if ignPol is not None:
            sTuple.append(self.zModifySettings(dst, "POP_IGNOREPOL", ignPol))
        if pos is not None:
            sTuple.append(self.zModifySettings(dst, "POP_POSITION", pos))
        if tiltx is not None:
            sTuple.append(self.zModifySettings(dst, "POP_TILTX", tiltx))
        if tilty is not None:
            sTuple.append(self.zModifySettings(dst, "POP_TILTY", tilty))
        return tuple(sTuple)
    
    def zSetPOPSettings(self, data=0, settingsFileName=None, start_surf=None, 
                        end_surf=None, field=None, wave=None, auto=None, 
                        beamType=None, paramN=((),()), pIrr=None, tPow=None, 
                        sampx=None, sampy=None, srcFile=None, widex=None, 
                        widey=None, fibComp=None, fibFile=None, fibType=None, 
                        fparamN=((),()), ignPol=None, pos=None, tiltx=None, 
                        tilty=None):
        """create and set a new settings file starting from the "reset" settings 
        state of the most basic lens in Zemax. 

        To modify an existing POP settings file, use ``zModifyPOPSettings()``. 
        Only those parameters with non-None or non-zero-length (in case of tuples)
        will be set.
        
        Parameters
        ----------
        data : integer
            0 = irradiance, 1 = phase
        settingsFileName : string, optional
            name to give to the settings file to be created. It must be the full 
            file name, including path and extension of settings file.
            If ``None``, then a CFG file with the name of the lens followed by 
            the string '_pyzdde_POP.CFG' will be created in the same directory 
            as the lens file and returned
        start_surf : integer, optional
            the starting surface
        end_surf : integer, optional
            the end surface
        field : integer, optional
            the field number
        wave : integer, optional
            the wavelength number
        auto : integer, optional
            simulates the pressing of the "auto" button which chooses appropriate
            X and Y beam widths based upon the sampling and other settings. Set
            it to 1
        beamType : integer (0...6), optional
            0 = Gaussian Waist; 1 = Gaussian Angle; 2 = Gaussian Size + Angle; 
            3 = Top Hat; 4 = File; 5 = DLL; 6 = Multimode.
        paramN : 2-tuple, optional
            sets beam parameter n, for example ((1, 4),(0.1, 0.5)) sets parameters 
            1 and 4 to 0.1 and 0.5 respectively. These parameter names and values 
            change depending upon the beam type setting. For example, for the 
            Gaussian Waist beam, n=1 for Waist X, 2 for Waist Y, 3 for Decenter X, 
            4 for Decenter Y, 5 for Aperture X, 6 for Aperture Y, 7 for Order X, 
            and 8 for Order Y
        pIrr : float, optional
            sets the normalization by peak irradiance. It is the initial beam peak 
            irradiance in power per area. It is an alternative to Total Power (tPow)
        tPow : float, optional
            sets the normalization by total beam power. It is the initial beam 
            total power. This is an alternative to Peak Irradiance (pIrr)
        sampx : integer (1...10), optional
            the X direction sampling. 1 for 32; 2 for 64; 3 for 128; 4 for 256; 
            5 for 512; 6 for 1024; 7 for 2048; 8 for 4096; 9 for 8192; 10 for 16384; 
        sampy : integer (1...10), optional
            the Y direction sampling. 1 for 32; 2 for 64; 3 for 128; 4 for 256; 
            5 for 512; 6 for 1024; 7 for 2048; 8 for 4096; 9 for 8192; 10 for 16384;
        srcFile : string, optional
            The file name if the starting beam is defined by a ZBF file, DLL, or 
            multimode file
        widex : float, optional
            the initial X direction width in lens units
        widey : float, optional
            the initial Y direction width in lens units
        fibComp : integer (1/0), optional
            use 1 to check the fiber coupling integral on, 0 to check it off
        fibFile : string, optional
            the file name if the fiber mode is defined by a ZBF or DLL
        fibType : string, optional
            use the same values as ``beamType`` above, except for multimode which 
            is not yet supported
        fparamN : 2-tuple, optional
            sets fiber parameter n, for example ((2,3),(0.5, 0.6)) sets parameters 
            2 and 3 to 0.5 and 0.6 respectively. See the hint for ``paramN``
        ignPol : integer (0/1), optional
            use 1 to ignore polarization, 0 to consider polarization
        pos : integer (0/1), optional
            use 0 for chief ray, 1 for surface vertex 
        tiltx : float, optional
            tilt about X in degrees
        tilty : float, optional
            tilt about Y in degrees

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just created 
            settings file
        
        Notes
        -----
        1. Further modifications of the settings file can be made using 
           ``ln.zModifySettings()`` or ``ln.zModifyPOPSettings()`` method
        2. The function creates settings file ending with '_pyzdde_POP.CFG'
           in order to prevent overwritting any existing settings file not
           created by pyzdde for POP.

        See Also
        -------- 
        zGetPOP()
        """
        # Create a settings file with "reset" settings
        global _pDir
        if data == 1:
            clean_cfg = 'RESET_SETTINGS_POP_PHASE.CFG'
        else:
            clean_cfg = 'RESET_SETTINGS_POP_IRR.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        
        if settingsFileName:
            dst = settingsFileName
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_POP.CFG'
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyPOPSettings(dst, start_surf, end_surf, field, wave, auto, 
                                    beamType, paramN, pIrr, tPow, sampx, sampy, 
                                    srcFile, widex, widey, fibComp, fibFile, 
                                    fibType, fparamN, ignPol, pos, tiltx, tilty)
            return dst       

    def zGetPupilMagnification(self):
        """Return the pupil magnification, which is the ratio of the exit-pupil
        diameter to the entrance pupil diameter.

        zGetPupilMagnification()->pupilMag

        Parameters
        ----------
        None

        Returns
        -------
        pupilMag    : (real value) The pupil magnification
        """
        _, _, ENPD, ENPP, EXPD, EXPP, _, _ = self.zGetPupil()
        return (EXPD/ENPD)

    def zGetSeidelAberration(self, which='wave', txtFile=None, keepFile=False):
        """Return the Seidel Aberration coefficients

        zGetSeidelAberration([which='wave', txtFile=None, keepFile=False]) -> sac

        Parameters
        ----------
        which           : (string, optional)
                          'wave' = Wavefront aberration coefficient (summary) is returned
                          'aber' = Seidel aberration coefficients (total) is returned
                          'both' = both Wavefront (summary) and Seidel aberration (total)
                                   coefficients are returned
        txtFile : (optional, string) If passed, the prescription file
                          will be named such. Pass a specific txtFileName if
                          you want to dump the file into a separate directory.
        keepFile        : (optional, bool) If false (default), the prescription
                          file will be deleted after use. If true, the file
                          will persist.
        Returns
        -------
        sac          : the Seidel aberration coefficients
                       if 'which' is 'wave', then a dictionary of Wavefront aberration
                       coefficient summary is returned.
                       if 'which' is 'aber', then a dictionary of Seidel total aberration
                       coefficient is returned
                       if 'which' is 'both', then a tuple of dictionaries containint Wavefront
                       aberration coefficients and Seidel aberration coefficients is returned.
        """
        if txtFile is not None:
            textFileName = txtFile
        else:
            cd = _os.path.dirname(_os.path.realpath(__file__))
            textFileName = _os.path.join(cd, "seidelAberrationFile.txt")
        ret = self.zGetTextFile(textFileName,'Sei', "None", 0)
        assert ret == 0
        recSystemData = self.zGetSystem() # Get the current system parameters
        numSurf = recSystemData[0]
        line_list = _readLinesFromFile(_openFile(textFileName))
        seidelAberrationCoefficients = {}         # Aberration Coefficients
        seidelWaveAberrationCoefficients = {}     # Wavefront Aberration Coefficients
        for line_num, line in enumerate(line_list):
            # Get the Seidel aberration coefficients
            sectionString1 = ("Seidel Aberration Coefficients:")
            if line.rstrip()== sectionString1:
                sac_keys_tmp = line_list[line_num + 2].rstrip()[7:] # remove "Surf" and "\n" from start and end
                sac_keys = sac_keys_tmp.split('    ')
                sac_vals = line_list[line_num + numSurf+3].split()[1:]
            # Get the Seidel Wavefront Aberration Coefficients (swac)
            sectionString2 = ("Wavefront Aberration Coefficient Summary:")
            if line.rstrip()== sectionString2:
                swac_keys01 = line_list[line_num + 2].split()     # names
                swac_vals01 = line_list[line_num + 3].split()[1:] # values
                swac_keys02 = line_list[line_num + 5].split()     # names
                swac_vals02 = line_list[line_num + 6].split()[1:] # values
                break
        else:
            raise Exception("Could not find section strings '{}'"
            " and '{}' in seidel aberrations file. "
            " \n\nPlease check if there is a mismatch in text encoding between"
            " Zemax and PyZDDE.".format(sectionString1, sectionString2))
        # Assert if the lengths of key-value lists are not equal
        assert len(sac_keys) == len(sac_vals)
        assert len(swac_keys01) == len(swac_vals01)
        assert len(swac_keys02) == len(swac_vals02)
        # Create the dictionary
        for k, v in zip(sac_keys, sac_vals):
            seidelAberrationCoefficients[k] = float(v)
        for k, v in zip(swac_keys01, swac_vals01):
            seidelWaveAberrationCoefficients[k] = float(v)
        for k, v in zip(swac_keys02, swac_vals02):
            seidelWaveAberrationCoefficients[k] = float(v)
        if not keepFile:
            _deleteFile(textFileName)
        if which == 'wave':
            return seidelWaveAberrationCoefficients
        elif which == 'aber':
            return seidelAberrationCoefficients
        elif which == 'both':
            return seidelWaveAberrationCoefficients, seidelAberrationCoefficients
        else:
            return None

    def zOptimize2(self, numCycle=1, algo=0, histLen=5, precision=1e-12,
                   minMF=1e-15, tMinCycles=5, tMaxCycles=None, timeout=None):
        """A wrapper around zOptimize() providing few control features.

        zOptimize2([numCycle, algo, histLen, precision, minMF,tMinCycles,
                  tMaxCycles])->(finalMerit, tCycles)

        Parameters
        ----------
        numCycles  : number of cycles per DDE call to optimization (default=1)
        algo       : 0=DLS, 1=Orthogonal descent (default=0)
        histLen    : length of the array of past merit functions returned from each
                     DDE call to zOptimize for determining steady state of merit
                     function values (default=5)
        precision  : minimum acceptable absolute difference between the merit-function
                     values in the array for steady state computation (default=1e-12)
        minMF      : minimum Merit Function following which to the optimization loop
                     is to be terminated even if a steady state hasn't reached.
                     This might be useful if a target merit function is desired.
        tMinCycles : total number of cycles to run optimization at the very least.
                     This is NOT the number of cycles per DDE call, but it is
                     calculated by multiplying the number of cycles per DDL optimize
                     call to the total number of DDE calls. (default=5)
        tMaxCycles : the maximum number of cycles after which the optimizaiton should
                     be terminated even if a steady state hasn't reached
        timeout    : timeout value (integer) in seconds used in each pass

        Returns
        -------
        finalMerit : (float) the final merit function.
        tCycles    : (integer) total number of cycles calculated by multiplying the
                     number of cycles per DDL optimize call to the total number of
                     DDE calls.

        Note
        ----
        `zOptimize2` basically calls `zOptimize` mutiple number of times in a loop.
        It can be useful if a large number of optimization cycles are required.
        """
        mfvList = [0.0]*histLen    # create a list of zeros
        count = 0
        mfvSettled = False
        finalMerit = 9e9
        tCycles = 0
        if not tMaxCycles:
            tMaxCycles = 2**31 - 1   # Largest plain positive integer value
        while not mfvSettled and (finalMerit > minMF) and (tCycles < tMaxCycles):
            finalMerit = self.zOptimize(numCycle, algo, timeout)
            self.zOptimize(-1,algo) # update all the operands in the MFE (not necessary?)
            if finalMerit > 8.9999e9: # optimization failure (Zemax returned 9.0E+009)
                break
            # populate mfvList in circular fashion
            mfvList[count % histLen] = finalMerit
            if (tCycles >= tMinCycles-1): # only after the minimum number of cycles are over,
                # test to see if the merit-function has settled down
                mfvList_shifted = mfvList[:-1]
                mfvList_shifted.append(mfvList[0])
                for i,j in zip(mfvList,mfvList_shifted):
                    if abs(i-j) >= precision:
                        break
                else:
                    mfvSettled = True
            count += 1
            tCycles = count*numCycle
        return (finalMerit, tCycles)


# ***************************************************************
#              IPYTHON NOTEBOOK UTILITY FUNCTIONS
# ***************************************************************
    def ipzCaptureWindowLQ(self, num=1, *args, **kwargs):
        """Capture graphic window from Zemax and display in IPython (Low Quality)

        ipzCaptureWindowLQ(num [, *args, **kwargs])-> displayGraphic

        Parameters
        ----------
        num: The graphic window to capture is indicated by the window number `num`.

        This function is useful for quickly capturing a graphic window, and
        embedding into a IPython notebook or QtConsole. The quality of JPG
        image is limited by the JPG export quality from Zemax.

        NOTE:
        ----
        In order to use this function, please copy the ZPL macros from
        PyZDDE\ZPLMacros to the macro directory where Zemax is expecting (i.e.
        as set in Zemax->Preference->Folders)

        For earlier versions (before 2010) please use ipzCaptureWindow().
        """
        global _global_IPLoad
        if _global_IPLoad:
            macroCode = "W{n}".format(n=str(num).zfill(2))
            dataPath = self.zGetPath()[0]
            imgPath = (r"{dp}\IMAFiles\{mc}_Win{n}.jpg"
                       .format(dp=dataPath, mc=macroCode, n=str(num).zfill(2)))
            if not self.zExecuteZPLMacro(macroCode):
                if _checkFileExist(imgPath):
                    _display(_Image(filename=imgPath))
                    _deleteFile(imgPath)
                else:
                    print("Timeout reached before image file was ready.")
                    print("The specified graphic window may not be open in ZEMAX!")
            else:
                print("ZPL Macro execution failed.\nZPL Macro path in PyZDDE is set to {}."
                      .format(self.macroPath))
                if not self.macroPath:
                    print("Use zSetMacroPath() to set the correct macro path.")
        else:
            print("Couldn't import IPython modules.")

    def ipzCaptureWindow2(self, *args, **kwargs):
        """Capture any analysis window from Zemax main window, using 3-letter
        analysis code. Same as ipzCaptureWindow().

        Note
        -----
        This function is now present only for backward compatibility.
        """
        return self.ipzCaptureWindow(*args, **kwargs)

    def ipzCaptureWindow(self, analysisType, percent=12, MFFtNum=0, blur=1,
                         gamma=0.35, settingsFile=None, flag=0, retArr=False,
                         wait=10):
        """Capture any analysis window from Zemax main window, using 3-letter
        analysis code.

        ipzCaptureWindow(analysisType [,percent=12,MFFtNum=0,blur=1, gamma=0.35,
                         settingsFile=None, flag=0, retArr=False])
                         -> displayGraphic

        Parameters
        ----------
        analysisType : string
                       3-letter button code for the type of analysis
        percent : float
                  percentage of the Zemax metafile to display (default=12).
                  Used for resizing the large metafile.
        MFFtNum : integer
                  type of metafile. 0 = Enhanced Metafile, 1 = Standard Metafile
        blur : float
               amount of blurring to use for antialiasing during resizing of
               metafile (default=1)
        gamma : float
                gamma for the PNG image (default = 0.35). Use a gamma value of
                around 0.9 for color surface plots.
        settingsFile : string
                           If a valid file name is used for the `settingsFile`,
                           ZEMAX will use or save the settings used to compute
                           the  metafile, depending upon the value of the flag
                           parameter.
        flag : integer
                0 = default settings used for the metafile graphic
                1 = settings provided in the settings file, if valid, else
                    default settings used
                2 = settings provided in the settings file, if valid, will be
                    used and the settings box for the requested feature will be
                    displayed. After the user makes any changes to the settings
                    the graphic will then be generated using the new settings.
        retArr : boolean
                whether to return the image as an array or not.
                If `False` (default), the image is embedded and no array is returned.
                If `True`, an numpy array is returned that may be plotted using Matpotlib.
        wait : time in sec
            time given to Zemax for the requested analysis to complete and produce
            a file.

        Returns
        -------
        None if `retArr` is False (default). The graphic is embedded into the notebook,
        else `pixel_array` (ndarray) if `retArr` is True.
        """
        global _global_IPLoad
        if _global_IPLoad:
            # Use the lens file path to store and process temporary images
            # tmpImgPath = self.zGetPath()[1]  # lens file path (default) ...
            # don't use the default lens path, as in earlier versions (before 2009)
            # of ZEMAX this path is in `C:\Program Files\Zemax\Samples`. Accessing
            # this folder to create the temporary file and then delete will most
            # likely not work due to permission issues.
            tmpImgPath = _os.path.dirname(self.zGetFile())  # directory of the lens file
            if MFFtNum==0:
                ext = 'EMF'
            else:
                ext = 'WMF'
            tmpMetaImgName = "{tip}\\TEMPGPX.{ext}".format(tip=tmpImgPath,ext=ext)
            tmpPngImgName = "{tip}\\TEMPGPX.png".format(tip=tmpImgPath)
            # Get the directory where PyZDDE (and thus `convert`) is located
            cd = _os.path.dirname(_os.path.realpath(__file__))
            # Create the ImageMagick command. At this time, we need two different
            # types of command because in Zemax:
            # 1. The Standard metafile export (as .WMF) seems to only work for
            #    layout analysis window, completely restricting its use
            # 2. The pen width setting in Zemax for the Enhanced metafile (as .EMF)
            #    is not functioning.
            if MFFtNum==0:
                imagickCmd = ("{cd}\convert {MetaImg} -flatten -blur {bl} "
                              "-resize {per}% -gamma {ga} {PngImg}"
                              .format(cd=cd,MetaImg=tmpMetaImgName,bl=blur,
                                      per=percent,ga=gamma,PngImg=tmpPngImgName))
            else:
                imagickCmd = ("{cd}\convert {MetaImg} -resize {per}% {PngImg}"
                              .format(cd=cd,MetaImg=tmpMetaImgName,per=percent,
                                                           PngImg=tmpPngImgName))
            # Get the metafile and display the image
            if not self.zGetMetaFile(tmpMetaImgName,analysisType,
                                     settingsFile,flag):
                if _checkFileExist(tmpMetaImgName, timeout=wait):
                    # Convert Metafile to PNG using ImageMagick's convert
                    startupinfo = _subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= _subprocess.STARTF_USESHOWWINDOW
                    _subprocess.Popen(args=imagickCmd, stdout=_subprocess.PIPE,
                                     startupinfo=startupinfo)
                    if _checkFileExist(tmpPngImgName,timeout=10):
                        _time.sleep(0.2)
                        if retArr:
                            if _global_mpl_img_load:
                                arr = _matimg.imread(tmpPngImgName, 'PNG')
                                _deleteFile(tmpMetaImgName)
                                _deleteFile(tmpPngImgName)
                                return arr
                            else:
                                print("Couldn't import Matplotlib")
                        else: # Display the image if not retArr
                            _display(_Image(filename=tmpPngImgName))
                            _deleteFile(tmpMetaImgName)
                            _deleteFile(tmpPngImgName)
                    else:
                        print("Timeout reached before PNG file was ready")
                else:
                    print("Timeout reached before Metafile file was ready")
            else:
                print("Metafile couldn't be created.")
        else:
                print("Couldn't import IPython modules.")

    def ipzGetTextWindow(self, analysisType, sln=0, eln=None, settingsFile=None,
                         flag=0, *args, **kwargs):
        """Print the text output of a Zemax analysis type into a IPython cell.

        ipzGetTextWindow(analysisType [,settingsFile, flag, *args, **kwargs])->
                                                                      textOutput

        Parameters
        ----------
        analysisType : 3 letter case-sensitive label that indicates the
                       type of the analysis to be performed. They are identical
                       to those used for the button bar in Zemax. The labels
                       are case sensitive. If no label is provided or recognized,
                       a standard raytrace will be generated.
        sln          : starting line number (integer) `default=0`
        eln          : ending line number (integer) `default=None`. If `None` all
                       lines in the file are printed.
        settingsFile : If a valid file name is used for the "settingsFile",
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

        Returns
        -------
        None (the contents of the text file is dumped into an IPython cell)
        """
        if not eln:
            eln = 1e10  # Set a very high number
        linePrintCount = 0
        global _global_IPLoad
        if _global_IPLoad:
            # Use the lens file path to store and process temporary images
            tmpTxtPath = self.zGetPath()[1]  # lens file path
            tmpTxtFile = "{ttp}\\TEMPTXT.txt".format(ttp=tmpTxtPath)
            if not self.zGetTextFile(tmpTxtFile,analysisType,settingsFile,flag):
                if _checkFileExist(tmpTxtFile):
                    for line in _getDecodedLineFromFile(_openFile(tmpTxtFile)):
                        if linePrintCount >= sln and linePrintCount <= eln:
                            print(line)  # print in the execution cell
                        linePrintCount += 1
                    _deleteFile(tmpTxtFile)
                else:
                    print("Text file of analysis window not created")
            else:
                print("GetTextFile didn't succeed")
        else:
            print("Couldn't import IPython modules.")

    def ipzGetFirst(self, pprint=True):
        """Print/ return first order data in human readable form

        Parameters
        ----------
        pprint : boolean. If True (default), the parameters are printed, else
                 a dictionary is returned.

        Returns
        -------
        Print or return dictionary containing first order data in human readable
        form that meant to be used in interactive environment.
        """
        firstData = self.zGetFirst()
        first = {}
        first['Effective focal length'] = firstData[0]
        first['Paraxial working F/#'] = firstData[1]
        first['Real working F/#'] = firstData[2]
        first['Paraxial image height'] =  firstData[3]
        first['Paraxial magnification'] = firstData[4]
        if pprint:
            _print_dict(first)
        else:
            return first

    def ipzGetPupil(self, pprint=True):
        """Print/ return pupil data in human readable form

        Parameters
        ----------
        pprint : boolean. If True (default), the parameters are printed, else
                 a dictionary is returned.

        Returns
        -------
        Print or return dictionary containing pupil information in human readable
        form that meant to be used in interactive environment.
        """
        pupilData = self.zGetPupil()
        pupil = {}
        apo_type = {0 : 'None', 1: 'Gaussian', 2 : 'Tangential/Cosine cubed'}
        pupil['Aperture Type'] = _system_aperture[pupilData[0]]
        if pupilData[0]==3: # if float by stop
            pupil['Value (stop surface semi-diameter)'] = pupilData[1]
        else:
            pupil['Value (system aperture)'] = pupilData[1]
        pupil['Entrance pupil diameter'] = pupilData[2]
        pupil['Entrance pupil position'] = pupilData[3]
        pupil['Exit pupil diameter'] = pupilData[4]
        pupil['Exit pupil position'] = pupilData[5]
        pupil['Apodization type'] = apo_type[pupilData[6]]
        pupil['Apodization factor'] = pupilData[7]
        if pprint:
            _print_dict(pupil)
        else:
            return pupil

    def ipzGetSystemAper(self, pprint=True):
        """Print or return system aperture data in human readable form

        Parameters
        ----------
        pprint : boolean. If True (default), the parameters are printed, else
                 a dictionary is returned.

        Returns
        -------
        Print or return dictionary containing system aperture information in
        human readable form that meant to be used in interactive environment.
        """
        sysAperData = self.zGetSystemAper()
        sysaper = {}
        sysaper['Aperture Type'] = _system_aperture[sysAperData[0]]
        sysaper['Stop surface'] = sysAperData[1]
        if sysAperData[0]==3: # if float by stop
            sysaper['Value (stop surface semi-diameter)'] = sysAperData[2]
        else:
            sysaper['Value (system aperture)'] = sysAperData[2]
        if pprint:
            _print_dict(sysaper)
        else:
            return sysaper

    def ipzGetSurfaceData(self, surfaceNumber, pprint=True):
        """Print or return basic (not all) surface data in human readable form

        Parameters
        ----------
        surfaceNumber : integer. Surface number
        pprint : boolean. If True (default), the parameters are printed, else
                 a dictionary is returned.

        Returns
        -------
        Print or return dictionary containing basic surface data (radius of curvature,
        thickness, glass, semi-diameter, and conic) in human readable form that
        meant to be used in interactive environment.
        """
        surfdata = {}
        surfdata['Radius of curvature'] = 1.0/self.zGetSurfaceData(surfaceNumber, 2)
        surfdata['Thickness'] = self.zGetSurfaceData(surfaceNumber, 3)
        surfdata['Glass'] = self.zGetSurfaceData(surfaceNumber, 4)
        surfdata['Semi-diameter'] = self.zGetSurfaceData(surfaceNumber, 5)
        surfdata['Conic'] = self.zGetSurfaceData(surfaceNumber, 6)
        if pprint:
            _print_dict(surfdata)
        else:
            return surfdata

    def ipzGetLDE(self):
        """Prints the sequential mode LDE data into the IPython cell

        Usage: ``ipzGetLDE()``

        Parameters
        ----------
        None

        Returns
        -------
        None

        Note
        ----
        Only works in sequential/hybrid mode. Can't retrieve NSC objects.
        """
        cd = _os.path.dirname(_os.path.realpath(__file__))
        textFileName = cd +"\\"+"prescriptionFile.txt"
        ret = self.zGetTextFile(textFileName,'Pre', "None", 0)
        assert ret == 0
        recSystemData = self.zGetSystem() # Get the current system parameters
        numSurf = recSystemData[0]
        line_list = _readLinesFromFile(_openFile(textFileName))
        for line_num, line in enumerate(line_list):
            sectionString = ("SURFACE DATA SUMMARY:") # to use re later
            if line.rstrip()== sectionString:
                for i in range(numSurf + 4): # 1 object surf + 3 extra lines before actual data
                    lde_line = line_list[line_num + i].rstrip()
                    print(lde_line)
                break
        else:
            raise Exception("Could not find string '{}' in Prescription file."
            " \n\nPlease check if there is a mismatch in text encoding between"
            " Zemax and PyZDDE, ``Surface Data`` is enabled in prescription"
            " file, and the mode is not pure NSC".format(sectionString))
        _deleteFile(textFileName)


# ***********************************************************************
#   OTHER HELPER FUNCTIONS THAT DO NOT REQUIRE A RUNNING ZEMAX SESSION
# ***********************************************************************
def numAper(aperConeAngle, rIndex=1.0):
    """Return the Numerical Aperture (NA) for the associated aperture cone angle

    numAper(aperConeAngle [, rIndex]) -> na

    Parameters
    ----------
    aperConeAngle : (float) aperture cone angle, in radians
    rIndex        : (float) refractive index of the medium

    Returns
    -------
    na : (float) Numerical Aperture
    """
    return rIndex*_math.sin(aperConeAngle)

def numAper2fnum(na, ri=1.0):
    """Convert numerical aperture (NA) to F-number

    Parameters
    ----------
    na : (float) Numerical aperture value
    ri : (float) Refractive index of the medium

    Returns
    -------
    fn : (float) F-number value
    """
    return 1.0/(2.0*_math.tan(_math.asin(na/ri)))

def fnum2numAper(fn, ri=1.0):
    """Convert F-number to numerical aperture (NA)

    Parameters
    ----------
    fn : (float) F-number value
    ri : (float) Refractive index of the medium

    Returns
    -------
    na : (float) Numerical aperture value
    """
    return ri*_math.sin(_math.atan(1.0/(2.0*fn)))

def fresnelNumber(r, z, wl=550e-6, approx=False):
    """calculate the fresnel number 
    
    Parameters
    ----------
    r : float
        radius of the aperture in units of length (usually mm)
    z : float
        distance of the observation plane from the aperture. this is equal 
        to the focal length of the lens for infinite conjugate, or the image 
        plane distance, in the same units of length as ``r``
    wl : float
        wavelength of light (default=550e-6 mm)
    approx : boolean
        if True, uses the approximate expression (default is False)
    
    Returns
    -------
    fN : float
        fresnel number
    
    Notes
    -----
    1. The Fresnel number is calculated based on a circular aperture or a 
       an unaberrated rotationally symmetric beam with finite extent [Zemax]_.
    2. From the Huygens-Fresnel principle perspective, the Fresnel number 
       represents the number of annular Fresnel zones in the aperture opening 
       [Wolf2011]_, or from the center of the beam to the edge in case of a
       propagating beam [Zemax]_. 

    References
    ----------
    .. [Zemax] Zemax manual

    .. [Born&Wolf2011] Principles of Optics, Born and Wolf, 2011
    """
    if approx:
        return (r**2)/(wl*z)
    else:
        return 2.0*(math.sqrt(z**2 + r**2) - z)/wl

# ***************************************************************************
# Helper functions to process data from ZEMAX DDE server. This is especially
# convenient for processing replies from Zemax for those function calls that
# a known data structure. These functions are mainly used intenally
# and may not be exposed directly.
# ***************************************************************************

def _regressLiteralType(x):
    """The function returns the literal with its proper type, such as int, float,
    or string from the input string x.
    Example:
        _regressLiteralType("1")->1
        _regressLiteralType("1.0")->1.0
        _regressLiteralType("1e-3")->0.001
        _regressLiteralType("YUV")->'YUV'
        _regressLiteralType("YO8")->'YO8'
    """
    try:
        float(x)  # Test for numeric or string
        lit = float(x) if set(['.','e','E']).intersection(x) else int(x)
    except ValueError:
        lit = str(x)
    return lit

def _checkFileExist(filename, mode='r', timeout=.25):
    """This function checks if a file exist

    If the file exist then it is ready to be read, written to, or deleted

    Parameters
    ----------
    filename : string
        filename with full path
    mode : string, optional 
        mode for opening file
    timeout : integer, 
        timeout in seconds for how long to wait before returning

    Returns
    -------
    status : bool
      True = file exist, and file operations are possible; 
      False = timeout reached
    """
    ti = _datetime.datetime.now()
    status = True
    while True:
        try:
            f = open(filename, mode)
        except IOError:
            timeDelta = _datetime.datetime.now() - ti
            if timeDelta.total_seconds() > timeout:
                status = False
                break
            else:
                _time.sleep(0.25)
        else:
            f.close()
            break
    return status

def _deleteFile(fileName, n=10):
    """Cleanly deletes a file. It takes n attempts to delete the file.

    If it can't delete the file in n attempts then it returns fail.

    _deleteFile(fileName [,n])->status

    Parameters
    ----------
    fileName  : file name of file to be deleted with full path
    n         : number of times to attempt before giving up

    Returns
    ------
    status   : True  = file deleting successful
               False = reached maximum number of attempts, without deleting file.

    Note
    ----
    It assumes that the file with filename actually exist and doesn't do any
    error checking on its existance. This is OK as this function is for internal
    use only.
    """
    status = False
    count = 0
    while not status and count < n:
        try:
            _os.remove(fileName)
        except OSError:
            count += 1
            _time.sleep(0.2)
        else:
            status = True
    return status

def _process_get_set_NSCProperty(code, reply):
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

def _process_get_set_Operand(column, reply):
    """Process reply for functions zGetOperand and zSetOperand"""
    rs = reply.rstrip()
    if column == 1:
        # ensure that it is a string ... as it is supposed to return the operand
        if isinstance(_regressLiteralType(rs), str):
            return str(rs)
        else:
            return -1
    elif column in (2,3):
        return int(float(rs))
    else:
        return float(rs)

def _process_get_set_Solve(reply):
    """Process reply for functions zGetSolve and zSetSolve"""
    reply = reply.rstrip()
    rs = reply.split(",")
    if 'BAD COMMAND' in rs:
        return -1
    else:
        return tuple([_regressLiteralType(x) for x in rs])

def _process_get_set_SystemProperty(code, reply):
    """Process reply for functions zGetSystemProperty and zSetSystemProperty"""
    # Convert reply to proper type
    if code in  (102,103, 104,105,106,107,108,109,110): # unexpected cases
        sysPropData = reply
    elif code in (16,17,23,40,41,42,43): # string
        sysPropData = reply.rstrip()    #str(reply)
    elif code in (11,13,24,53,54,55,56,60,61,62,63,71,72,73,77,78): # floats
        sysPropData = float(reply)
    else:
        sysPropData = int(float(reply))      # integer
    return sysPropData

def _process_get_set_Tol(operandNumber,reply):
    """Process reply for functions zGetTol and zSetTol"""
    rs = reply.rsplit(",")
    tolType = [rs[0]]
    tolParam = [float(e) if i in (2,3) else int(float(e))
                                 for i,e in enumerate(rs[1:])]
    toleranceData = tuple(tolType + tolParam)
    return toleranceData

def _print_dict(data):
    """Helper function to print a dictionary so that the key and value are
    arranged into nice rows and columns
    """
    leftColMaxWidth = max(_imap(len, data))
    for key, value in data.items():
        print("{}: {}".format(key.ljust(leftColMaxWidth + 1), value))

def _openFile(fileName):
    """opens the file in the appropriate mode and returns the file object

    Parameters
    ----------
    fileName (string) : name of the file to open

    Returns
    -------
    f (file object)

    Note
    ----
    This is just a wrapper around the open function.
    It is the responsibility of the calling function to close the file by
    calling the close() method of the file object. Alternatively use either use
    a with/as context to close automatically or use _readLinesFromFile() or
    _getDecodedLineFromFile() that uses a with context manager to handle
    exceptions and file close.
    """
    global _global_use_unicode_text
    if _global_use_unicode_text:
        f = open(fileName, u'rb')
    else:
        f = open(fileName, 'r')
    return f

def _getDecodedLineFromFile(fileObj):
    """generator function; yields a decoded (ascii/Unicode) line
    The file is automatically closed when after reading the file or if any
    exception occurs while reading the file.
    """
    global _global_pyver3
    global _global_use_unicode_text
    global _global_in_IPython_env

    # I am not exactly sure why there is a difference in behavior
    # between IPython environment and normal Python shell, but it is there!
    if _global_in_IPython_env:
        unicode_type = 'utf-16-le'
    else:
        unicode_type = 'utf-16'

    if _global_use_unicode_text:
        fenc = _codecs.EncodedFile(fileObj, unicode_type)
        with fenc as f:
            for line in f:
                decodedLine = line.decode(unicode_type)
                yield decodedLine.rstrip()
    else: # ascii
        with fileObj as f:
            for line in f:
                if _global_pyver3: # ascii and Python 3.x
                    yield line.rstrip()
                else:      # ascii and Python 2.x
                    try:
                        decodedLine = line.decode('raw-unicode-escape')
                    except:
                        decodedLine = line.decode('ascii', 'replace')
                    yield decodedLine.rstrip()

def _readLinesFromFile(fileObj):
    """returns a list of lines (as unicode literals) in the file

    This function emulates the functionality of readlines() method of file
    objects. The caller doesn't have to explicitly close the file as it is
    handled in _getDecodedLineFromFile() function.

    Parameters
    ----------
    fileObj : file object returned by open() method

    Returns
    -------
    lines (list) : list of lines (as unicode literals with u'string' notation)
                   from the file
    """
    lines = list(_getDecodedLineFromFile(fileObj))
    return lines

def _getFirstLineOfInterest(line_list, pattern, patAtStart=True):
    """returns the line number (index in the list of lines) that matches the 
    regex pattern. 

    This function can be used to return the starting line of the data-of-interest, 
    identified by the regex pattern, from a list of lines.
    
    Parameters
    ----------
    line_list : list
        list of lines in the file returned by ``_readLinesFromFile()``
    pattern : string
        regex pattern that should be used to identify the line of interest
    patAtStart : bool
        if ``True``, match pattern at the beginning of line string (default)
    
    Returns
    -------
    line_number : integer
        line_number/ index in the list where the ``pattern`` first matched. 
        If no match could be found, the function returns ``None``

    Notes
    ----- 
    If it is known that the pattern will be matched at the beginning, then
    letting ``patAtStart==True`` is more efficient.
    """
    pat = _re.compile(pattern) if patAtStart else _re.compile('.*'+pattern)
    for line_num, line in enumerate(line_list):
        if _re.match(pat, line.strip()):
            return line_num

def _get2DList(line_list, start_line, number_of_lines):
    """returns a 2D list of data read between ``start_line`` and 
    ``start_line + number_of_lines`` of a list 
    
    Parameters
    ----------
    line_list : list
        list of lines read from a file using ``_readLinesFromFile()``
    start_line : integer
        index of line_list 
    number_of_lines : integer
        number of lines to read (number of lines which contain the 2D data)

    Returns
    ------- 
    data : list
        data is a 2-D list of float type data read from the lines in line_list
    """
    data = []
    end_Line = start_line + number_of_lines
    for lineNum, row in enumerate(line_list):
        if start_line <= lineNum <= end_Line:
            data.append([float(i) for i in row.split('\t')])
    return data
#
#
if __name__ == "__main__":
    print("Please import this module as 'import pyzdde.zdde as pyz' ")
    _sys.exit(0)

