# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        zdde.py
# Purpose:     Python based DDE link with ZEMAX server, similar to Matlab based
#              MZDDE toolbox.
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    2.0.3
#-------------------------------------------------------------------------------
"""PyZDDE, which is a toolbox written in Python, is used for communicating
with ZEMAX using the Microsoft's Dynamic Data Exchange (DDE) messaging
protocol. The docstring examples in the functions assume that PyZDDE is
imported as ``import pyzdde.zdde as pyz`` and a PyZDDE communication object
is then created as ``ln = pyz.createLink()`` or ``ln = pyz.PyZDDE();
ln.zDDEInit()``.
"""
from __future__ import division
from __future__ import print_function
import sys as _sys
#import struct as _struct
import os as _os
import collections as _co
import subprocess as _subprocess
import math as _math
import time as _time
import datetime as _datetime
import re as _re
import shutil as _shutil
import warnings as _warnings
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

# The first module to import that is not one of the standard modules MUST
# be the config module as it sets up the different global and settings variables
_currDir = _os.path.dirname(_os.path.realpath(__file__))
_pDir = _os.path.split(_currDir)[0]

settings_file = _os.path.join(_currDir, "settings.ini")
if not _os.path.isfile(settings_file):
    src = _os.path.join(_currDir, "settings.ini-dist")
    _shutil.copy(src, settings_file)

import pyzdde.config as _config
_global_pyver3 = _config._global_pyver3
_global_use_unicode_text = _config._global_use_unicode_text
imageMagickSettings = _config.getImageMagickSettings()
_global_use_installed_imageMagick = imageMagickSettings[0]
_global_imageMagick_dir = imageMagickSettings[1]

# DDEML communication module
_global_ddeclient_load = False # True if module could be loaded.
try:
  import pyzdde.ddeclient as _dde
  _global_ddeclient_load = True
except ImportError:
  # System may not be windows; only provide functions that do not use zemax
  print("DDE client couldn't be loaded. All functions prefixed with"
        " \"z\" or \"ipz\" may not work.")

# Python 2/ Python 3 differential imports
if _global_pyver3:
   _izip = zip
   _imap = map
   xrange = range
   import tkinter as _tk
   import tkinter.messagebox as _MessageBox
else:
    from itertools import izip as _izip, imap as _imap
    import Tkinter as _tk
    import tkMessageBox as _MessageBox
    
# Pyzdde local module imports
import pyzdde.zcodes.zemaxbuttons as zb
import pyzdde.zcodes.zemaxoperands as zo
import pyzdde.utils.pyzddeutils as _putils
import pyzdde.zfileutils as _zfu

#%% Constants
_DEBUG_PRINT_LEVEL = 0 # 0 = No debug prints, but allow all essential prints
                       # 1 to 2 levels of debug print, 2 = print all

_MAX_PARALLEL_CONV = 2  # Max no of simul. conversations possible with Zemax
_system_aperture = {0 : 'EPD',
                    1 : 'Image space F/#',
                    2 : 'Object space NA',
                    3 : 'Float by stop',
                    4 : 'Paraxial working F/#',
                    5 : 'Object cone angle'}

macheps = _sys.float_info.epsilon  # machine epsilon

#%% Helper function for debugging
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


#%% Module methods

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

# decorator for automatically push and refresh to and from LDE (Experimental)
def autopushandrefresh(func): 
    def wrapped(self, *args, **kwargs):
        if self.apr: # if automatic push refresh is True
            if (args[0].startswith('Get') or
                args[0].startswith('Set') or
                args[0].startswith('Insert') or 
                args[0].startswith('Delete')):
                self._conversation.Request('GetRefresh')
            reply = func(self, *args, **kwargs)
            if (args[0].startswith('Set') or 
                args[0].startswith('Insert') or 
                args[0].startswith('Delete')):
                self._conversation.Request('PushLens,1')
        else:
            reply = func(self, *args, **kwargs)
        return reply
    return wrapped 


_global_dde_linkObj = {}

def createLink(apr=False):
    """Create a DDE communication link with Zemax

    Usage: ``import pyzdde.zdde as pyz; ln = pyz.createLink()``

    Helper function to create, initialize and return a PyZDDE communication
    object.

    Parameters
    ----------
    apr : bool 
        if `True`, automatically push and refresh lens to and from LDE to DDE 

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
        link = PyZDDE(apr=apr)
        status = link.zDDEInit()
        if not status:
            _global_dde_linkObj[link] = link._appName  # This can be something more useful later
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
        If a specific link object is not given, all existing links are
        closed.

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

def setImageMagickSettings(use_installed_ImageMagick, imageMagick_dir=None):
    """Set the use-flag and imageMagick installation directory settings

    Parameters
    ----------
    use_installed_ImageMagick : bool
        boolean flag to indicate whether to use installed version
        of ImageMagick (``True``) or not (``False``)
    imageMagick_dir : string, optional
        full path to the installation directory. For example:
        ``C:\\Program Files\\ImageMagick-6.8.9-Q8``

    Returns
    -------
    imageMagick_settings : tuple
        updated imageMagick settings
    """
    global _global_use_installed_imageMagick
    global _global_imageMagick_dir
    if not isinstance(use_installed_ImageMagick, bool):
        raise ValueError("Expecting bool")
    if imageMagick_dir and not _os.path.isdir(imageMagick_dir):
        raise ValueError("Expecting valid directory or None")
    if imageMagick_dir and not _os.path.isfile(_os.path.join(imageMagick_dir,
                                              'convert.exe')):
        raise ValueError("Couldn't find program convert.exe in the path!")
    _config.setImageMagickSettings(use_installed_ImageMagick, imageMagick_dir)
    imageMagickSettings = _config.getImageMagickSettings()
    _global_use_installed_imageMagick = imageMagickSettings[0]
    _global_imageMagick_dir = imageMagickSettings[1]
    return (_global_use_installed_imageMagick, _global_imageMagick_dir)

def getImageMagickSettings():
    """Return the use-flag and imageMagick installation directory settings

    Parameters
    ----------
    None

    Returns
    -------
    use_flag : bool
        if ``True``, then PyZDDE uses the installed version of ImageMagick
        software. If ``False``, then the version of ImageMagick that comes
        with PyZDDE will be used.
    imageMagick_dir : string
        ImageMagick installation directory.
    """
    return _config.getImageMagickSettings()

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

#%% PyZDDE class

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

    version = '2.0.3'
    
    # Other class variables
    # Surface data codes for getting and setting surface data
    SDAT_TYPE = 0   # Surface type name
    SDAT_COMMENT = 1  # Comment
    SDAT_CURV = 2   # Curvature
    SDAT_THICK = 3  # Thickness
    SDAT_GLASS = 4   # Glass
    SDAT_SEMIDIA = 5  # Semi-Diameter
    SDAT_CONIC = 6  # Conic
    SDAT_COAT = 7  # Coating
    SDAT_TCE = 8  # Thermal Coefficient of Expansion (TCE)
    SDAT_UD_DLL = 9  # User-defined DLL
    SDAT_IGNORE_S_FLAG = 20  # Ignore surface flag
    SDAT_TILT_DCNTR_ORD_BEFORE = 51  # Before tilt and decenter order
    SDAT_DCNTR_X_BEFORE = 52  # Before decenter x 
    SDAT_DCNTR_Y_BEFORE = 53  # Before decenter y
    SDAT_TILT_X_BEFORE = 54  # Before tilt x 
    SDAT_TILT_Y_BEFORE = 55  # Before tilt y
    SDAT_TILT_Z_BEFORE = 56  # Before tilt z
    SDAT_TILT_DCNTR_STAT_AFTER = 60  # After status 
    SDAT_TILT_DCNTR_ORD_AFTER = 61  # After tilt and decenter order
    SDAT_DCNTR_X_AFTER = 62  # After decenter x
    SDAT_DCNTR_Y_AFTER = 63  # After decenter y
    SDAT_TILT_X_AFTER = 64  # After tilt x
    SDAT_TILT_Y_AFTER = 65  # After tilt y
    SDAT_TILT_Z_AFTER = 66  # After tilt z
    SDAT_USE_LAYER_MULTI_INDEX = 70  # Use Layer multipliers and index offsets
    SDAT_LAYER_MULTI_VAL = 71  # Layer multiplier value
    SDAT_LAYER_MULTI_STAT = 72 # Layer multiplier status
    SDAT_LAYER_INDEX_OFFSET_VAL = 73  # Layer index offset value
    SDAT_LAYER_INDEX_OFFSET_STAT = 74  # Layer index offset status
    SDAT_LAYER_EXTINCT_OFFSET_VAL = 75   # Layer extinction offset value
    SDAT_LAYER_EXTINCT_OFFSET_STAT = 76  # Layer extinction offset status
    # Surface parameter codes for getting and setting solves
    SOLVE_SPAR_CURV = 0   # Curvature
    SOLVE_SPAR_THICK = 1  # Thickness
    SOLVE_SPAR_GLASS = 2   # Glass
    SOLVE_SPAR_SEMIDIA = 3  # Semi-Diameter
    SOLVE_SPAR_CONIC = 4  # Conic
    SOLVE_SPAR_PAR0 = 17  # Parameter 0
    SOLVE_SPAR_PAR1 = 5  # Parameter 1
    SOLVE_SPAR_PAR2 = 6  # Parameter 2  
    SOLVE_SPAR_PAR3 = 7  # Parameter 3
    SOLVE_SPAR_PAR4 = 8  # Parameter 4
    SOLVE_SPAR_PAR5 = 9  # Parameter 5
    SOLVE_SPAR_PAR6 = 10  # Parameter 6
    SOLVE_SPAR_PAR7 = 11  # Parameter 7
    SOLVE_SPAR_PAR8 = 12  # Parameter 8
    SOLVE_SPAR_PAR9 = 13  # Parameter 9
    SOLVE_SPAR_PAR10 = 14  # Parameter 10
    SOLVE_SPAR_PAR11 = 15  # Parameter 11
    SOLVE_SPAR_PAR12 = 16  # Parameter 12
    # Solve type code for use with get/set solve function
    SOLVE_CURV_FIXED = 0 # Solve on curvature; fixed
    SOLVE_CURV_VAR = 1 # Solve on curvature; variable (V)
    SOLVE_CURV_MR_ANG = 2 # Solve on curvature; marginal ray angle (M) 
    SOLVE_CURV_CR_ANG = 3 # Solve on curvature; chief ray angle (C)
    SOLVE_CURV_PICKUP = 4 # Solve on curvature; pickup (P)
    SOLVE_CURV_MR_NORM = 5 # Solve on curvature; marginal ray normal (N)
    SOLVE_CURV_CR_NORM = 6 # Solve on curvature; chief ray normal (N)
    SOLVE_CURV_APLAN = 7 # Solve on curvature; aplanatic (A)
    SOLVE_CURV_ELE_POWER = 8 # Solve on curvature; element power (X)
    SOLVE_CURV_CON_SURF = 9 # Solve on curvature; concentric with surface (S)
    SOLVE_CURV_CON_RADIUS = 10 # Solve on curvature; concentric with radius (R)
    SOLVE_CURV_FNUM = 11 # Solve on curvature; f/# (F)
    SOLVE_CURV_ZPL = 12 # Solve on curvature; zpl macro (Z)
    SOLVE_THICK_FIXED = 0 # Solve on thickness; fixed
    SOLVE_THICK_VAR = 1 # Solve on thickness; variable (V)
    SOLVE_THICK_MR_HGT = 2 # Solve on thickness; marginal ray height (M)
    SOLVE_THICK_CR_HGT = 3 # Solve on thickness; chief ray height (C)
    SOLVE_THICK_EDGE_THICK = 4 # Solve on thickness; edge thickness (E)
    SOLVE_THICK_PICKUP = 5 # Solve on thickness; pickup (P)
    SOLVE_THICK_OPD = 6 # Solve on thickness; optical path difference (O)
    SOLVE_THICK_POS = 7 # Solve on thickness; position (T)
    SOLVE_THICK_COMPENSATE = 8 # Solve on thickness; compensator (S)
    SOLVE_THICK_CNTR_CURV = 9 # Solve on thickness; center of curvature (X)
    SOLVE_THICK_PUPIL_POS = 10 # Solve on thickness; pupil position (U)
    SOLVE_THICK_ZPL = 11 # Solve on thickness; zpl macro (Z)
    SOLVE_GLASS_FIXED = 0 # Solve on glass; fixed
    SOLVE_GLASS_MODEL = 1 # Solve on glass; model
    SOLVE_GLASS_PICKUP = 2 # Solve on glass; pickup (P)
    SOLVE_GLASS_SUBS = 3 # Solve on glass; substitute (S)
    SOLVE_GLASS_OFFSET = 4 # Solve on glass; offset (O)
    SOLVE_SEMIDIA_AUTO = 0 # Solve on semi-diameter; automatic
    SOLVE_SEMIDIA_FIXED = 1 # Solve on semi-diameter; fixed (U)
    SOLVE_SEMIDIA_PICKUP = 2 # Solve on semi-diameter; pickup (P)
    SOLVE_SEMIDIA_MAX = 3 # Solve on semi-diameter; maximum (M)
    SOLVE_SEMIDIA_ZPL = 4 # Solve on semi-diameter; zpl macro (Z)
    SOLVE_CONIC_FIXED = 0 # Solve on conic; fixed
    SOLVE_CONIC_VAR = 1 # Solve on conic; variable (V)
    SOLVE_CONIC_PICKUP = 2 # Solve on conic; pickup (P)
    SOLVE_CONIC_ZPL = 3 # Solve on conic; zpl macro (Z)
    SOLVE_PAR0_FIXED = 0 # Solve on parameter 0; fixed
    SOLVE_PAR0_VAR = 1 # Solve on parameter 0; variable (V)
    SOLVE_PAR0_PICKUP = 2 # Solve on parameter 0; pickup (P)
    SOLVE_PARn_FIXED = 0 # Solve on parameter n (b/w 1 - 12); fixed
    SOLVE_PARn_VAR = 1 # Solve on parameter n (b/w 1 - 12); variable (V)
    SOLVE_PARn_PICKUP = 2 # Solve on parameter n (b/w 1 - 12); pickup (P)
    SOLVE_PARn_CR = 3 # Solve on parameter n (b/w 1 - 12); chief-ray (C)
    SOLVE_PARn_ZPL = 4 # Solve on parameter n (b/w 1 - 12); zpl macro (Z)
    SOLVE_EDATA_FIXED = 0 # Solve on extra data values; fixed
    SOLVE_EDATA_VAR = 1 # Solve on extra data values; variable (V)
    SOLVE_EDATA_PICKUP = 2 # Solve on extra data values; pickup (P)
    SOLVE_EDATA_ZPL = 3 # Solve on extra data values; zpl macro (Z)
    # Object parameter codes for NSC solve
    NSCSOLVE_OPAR_XPOS = -1
    NSCSOLVE_OPAR_YPOS = -2
    NSCSOLVE_OPAR_ZPOS = -3
    NSCSOLVE_OPAR_XTILT = -4
    NSCSOLVE_OPAR_YTILT = -5
    NSCSOLVE_OPAR_ZTILT = -6
    ANA_POP_SAMPLE_32 = 1
    # Sampling codes for POP analysis
    ANA_POP_SAMPLE_64 = 2 
    ANA_POP_SAMPLE_128 = 3 
    ANA_POP_SAMPLE_256 = 4 
    ANA_POP_SAMPLE_512 = 5 
    ANA_POP_SAMPLE_1024 = 6 
    ANA_POP_SAMPLE_2048 = 7 
    ANA_POP_SAMPLE_4096 = 8
    ANA_POP_SAMPLE_8192 = 9 
    ANA_POP_SAMPLE_16384 = 10
    # Sampling codes for PSF/MTF analysis
    ANA_PSF_SAMPLE_32x32 = 1
    ANA_PSF_SAMPLE_64x64 = 2 
    ANA_PSF_SAMPLE_128x128 = 3 
    ANA_PSF_SAMPLE_256x256 = 4 
    ANA_PSF_SAMPLE_512x512 = 5 
    ANA_PSF_SAMPLE_1024x1024 = 6 
    ANA_PSF_SAMPLE_2048x2048 = 7 
    ANA_PSF_SAMPLE_4096x4096 = 8
    ANA_PSF_SAMPLE_8192x8192 = 9 
    ANA_PSF_SAMPLE_16384x16384 = 10

    def __init__(self, apr=False):
        """Creates an instance of PyZDDE class

        Usage: ``ln = pyz.PyZDDE()``

        Parameters
        ----------
        apr : bool 
            if `True`, automatically push and refresh lens to and from LDE to DDE

        Returns
        -------
        ln : PyZDDE object

        Notes
        -----
        1. Following the creation of PyZDDE object, initiate the
           communication channel as ``ln.zDDEInit()``
        2. Consider using the module level function ``pyz.createLink()`` to
           create and initiate a DDE channel instead of ``ln = pyz.PyZDDE();
           ln.zDDEInit()``

        See Also
        --------
        createLink()
        """
        PyZDDE.__chNum += 1   # increment channel count
        self._appName = _getAppName(PyZDDE.__appNameDict) or '' # wicked :-)
        self._appNum = PyZDDE.__chNum # unique & immutable identity of each instance
        self._connection = False  # 1/0 depending on successful connection or not
        self._macroPath = None    # variable to store macro path
        self._filesCreated = set()   # .cfg & other files to be cleaned at session end
        self._apr = apr

    def __repr__(self):
        return ("PyZDDE(appName=%r, appNum=%r, connection=%r, macroPath=%r)" %
                (self._appName, self._appNum, self._connection, self._macroPath))

    def __hash__(self):
        # for storing in internal dictionary
        return hash(self._appNum)

    def __eq__(self, other):
        return (self._appNum == other._appNum)

    @property
    def apr(self):
        return self._apr

    @apr.setter
    def apr(self, val):
        self._apr = val

    

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
        _debugPrint(1,"appName = " + self._appName)
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
        self._conversation = _dde.CreateConversation(PyZDDE.__server)
        _debugPrint(2, "PyZDDE.converstation = " + str(self._conversation))
        try:
            self._conversation.ConnectTo(self._appName," ")
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
            self._connection = True
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
        This bounded method provides a quick alternative way to close link
        rather than calling the module function ``pyz.closeLink()``.

        See Also
        --------
        zDDEClose() :
            PyZDDE instance method to close a link.
            Use this method (as ``ln.zDDEClose()``) if the link was
            created as ``ln = pyz.PyZDDE(); ln.zDDEInit()``
        closeLink() :
            A moudle level function to close a link.
            Use this method (as ``pyz.closeLink(ln)``) or ``ln.close()``
            if the link was created as ``ln = pyz.createLink()``
        """
        return closeLink(self)

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
        Use this bounded method to close link if the link was created using
        the idiom ``ln = pyz.PyZDDE(); ln.zDDEInit()``. If however, the
        link was created using ``ln = pyz.createLink()``, use either
        ``pyz.closeLink()`` or ``ln.close()``.
        """
        if PyZDDE.__server and not PyZDDE.__liveCh:
            PyZDDE.__server.Shutdown(self._conversation)
            PyZDDE.__server = 0
            _debugPrint(2,"server shutdown as ZEMAX is not running!")
        elif PyZDDE.__server and self._connection and PyZDDE.__liveCh == 1:
            PyZDDE.__server.Shutdown(self._conversation)
            self._connection = False
            PyZDDE.__appNameDict[self._appName] = False # make the name available
            _deleteFilesCreatedDuringSession(self)
            self._appName = ''
            PyZDDE.__liveCh -= 1  # This will become zero now. (reset)
            PyZDDE.__server = 0   # previous server obj should be garbage collected
            _debugPrint(2,"server shutdown")
        elif self._connection:  # if additional channels were successfully created.
            PyZDDE.__server.Shutdown(self._conversation)
            self._connection = False
            PyZDDE.__appNameDict[self._appName] = False # make the name available
            _deleteFilesCreatedDuringSession(self)
            self._appName = ''
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
            timeout value in seconds (if float is given, it is rounded to
            integer)

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
        self._conversation.SetDDETimeout(round(time))
        return self._conversation.GetDDETimeout()

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
        return self._conversation.GetDDETimeout()

    @autopushandrefresh
    def _sendDDEcommand(self, cmd, timeout=None):
        """Method to send command to DDE client
        """
        global _global_pyver3
        reply = self._conversation.Request(cmd, timeout)
        if _global_pyver3:
            reply = reply.decode('ascii').rstrip()
        return reply

    def __del__(self):
        """Destructor"""
        _debugPrint(2,"Destructor called")
        self.zDDEClose()

    # ****************************************************************
    #               ZEMAX DATA ITEM BASED METHODS
    # ****************************************************************
    def zCloseUDOData(self, bufferCode):
        """Close the User Defined Operand buffer allowing optimizer to
        proceed

        Parameters
        ----------
        bufferCode : integer
            buffercode is an integer value provided by Zemax to the client
            that uniquely identifies the correct lens.

        Returns
        -------
        retVal : ?

        See Also
        --------
         zGetUDOSystem(), zSetUDOItem()
        """
        return int(self._sendDDEcommand("CloseUDOData,{:d}".format(bufferCode)))

    def zDeleteConfig(self, number):
        """Deletes an existing configuration (column) in the multi-
        configuration editor

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
        After deleting the configuration, all succeeding configurations
        are re-numbered.

        See Also
        --------
        zInsertConfig()

        zDeleteMCO() :
            (TIP) use zDeleteMCO to delete a row/operand
        """
        return int(self._sendDDEcommand("DeleteConfig,{:d}".format(number)))

    def zDeleteMCO(self, operNum):
        """Deletes an existing operand (row) in the multi-configuration
        editor

        Parameters
        ----------
        operNum : integer
            operand number (row in the MCE) to delete

        Returns
        -------
        newNumberOfOperands : integer
            new number of operands

        Notes
        -----
        After deleting the row, all succeeding rows (operands) are
        re-numbered.

        See Also
        --------
        zInsertMCO()
        zDeleteConfig() :
            (TIP) Use zDeleteConfig() to delete a column/configuration.
        """
        return int(self._sendDDEcommand("DeleteMCO,"+str(operNum)))

    def zDeleteMFO(self, operand):
        """Deletes an optimization operand (row) in the merit function
        editor

        Parameters
        ----------
        operand : integer
            Operand (row) number (- 1 <= operand <= number_of_operands)

        Returns
        -------
        newNumOfOperands : integer
            the new number of operands

        See Also
        --------
        zInsertMFO()
        """
        return int(self._sendDDEcommand("DeleteMFO,{:d}".format(operand)))

    def zDeleteObject(self, surfNum, objNum):
        """Deletes the NSC object identified by the ``objNum`` and
        the surface identified by ``surfNum``

        Parameters
        ----------
        surfNum : integer
            surface number of Non-Sequential Component surface
        objNum : integer
            object number in the NSC editor

        Returns
        -------
        status : integer (0 or -1)
            0 if successful, -1 if it failed

        Notes
        -----
        1. The ``surfNum`` is 1 if the lens is purely NSC mode.
        2. If no more objects are present it simply returns 0.

        See Also
        --------
        zInsertObject()
        """
        cmd = "DeleteObject,{:d},{:d}".format(surfNum,objNum)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            return -1
        else:
            return int(float(rs))

    def zDeleteSurface(self, surfNum):
        """Deletes an existing surface identified by ``surfNum``

        Parameters
        ----------
        surfNum : integer
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
        cmd = "DeleteSurface,{:d}".format(surfNum)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply))

    def zExportCAD(self, fileName, fileType=1, numSpline=32, firstSurf=1,
                   lastSurf=-1, raysLayer=1, lensLayer=0, exportDummy=0,
                   useSolids=1, rayPattern=0, numRays=0, wave=0, field=0,
                   delVignett=1, dummyThick=1.00, split=0, scatter=0,
                   usePol=0, config=0):
        """Export lens data in IGES/STEP/SAT format for import into CAD
        programs

        Parameters
        ----------
        fileName : string
            filename including extension (including full path is
            recommended)
        fileType : integer (0, 1, 2 or 3)
            0 = IGES; 1 = STEP (default); 2 = SAT; 3 = STL
        numSpline : integer
            number of spline segments to use (default = 32)
        firstSurf : integer
            the first surface to export; the first object to export
            (in NSC mode)
        lastSurf : integer
            the last surface to export; the last object to export
            (in NSC mode)
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
            split rays from NSC sources? 1 = split sources;
            0 (default) = no
        scatter : integer (0 or 1)
            scatter rays from NSC sources? 1 = Scatter; 0 (deafult) = no
        usePol : integer (0 or 1)
            use polarization when tracing NSC rays? 1 = use polarization;
            0 (default) no. Note that polarization is automatically
            selected if ``split`` is ``1``.
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
        1. Exporting lens data data may take longer than the timeout
           interval of the DDE communication. Zemax spwans an independent
           thread to process this request. Once the thread is launched,
           Zemax returns "Exporting filename". However, the export may
           take much longer. To verify the completion of export and the
           readiness of the file, use ``zExportCheck()``, which returns
           ``1`` as long as the export is in process, and ``0`` once
           completed. Generally, ``zExportCheck()`` should be placed in
           a loop, which executes until a ``0`` is returned.

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

        2. Zemax cannot export some NSC objects (e.g. slide). The
           unexportable objects are ignored.

        References
        ----------
        For a detailed exposition on the configuration settings,
        see "Export IGES/SAT.STEP Solid" in the Zemax manual [Zemax]_.
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
        """Returns the surface that has the integer label associated with
        the it.

        Parameters
        ----------
        label : integer
            label associated with a surface

        Returns
        -------
        surfNum : integer
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

    def zGetAperture(self, surf):
        """Get the surface aperture data for a given surface

        Parameters
        ----------
        surf : integer
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
            min radius(ca); min radius(co); width of arm(s); x-half width
            (ra); x-half width(ro); x-half width(ea); x-half width(eo)
        aMax : float
            max radius(ca); max radius(co); number of arm(s); y-half
            width(ra); y-half width(ro); y-half width(ea); y-half width(eo)
        xDecenter : float
            amount of decenter in x from current optical axis (lens units)
        yDecenter : float
            amount of decenter in y from current optical axis (lens units)
        apertureFile : string
            a text file with .UDA extention.

        References
        ----------
        The following sections from the Zemax manual should be referred
        for details [Zemax]_:

        1. "Aperture type and other aperture controls" for details on
           aperture
        2. "User defined apertures and obscurations" for more on UDA
           extension

        See Also
        --------
        zGetSystemAper() :
            For system aperture instead of the aperture of surface.
        zSetAperture()
        """
        reply = self._sendDDEcommand("GetAperture," + str(surf))
        rs = reply.split(',')
        apertureInfo = [int(rs[i]) if i==5 else float(rs[i])
                                             for i in range(len(rs[:-1]))]
        apertureInfo.append(rs[-1].rstrip()) # append the test file (string)
        ainfo = _co.namedtuple('ApertureInfo', ['aType', 'aMin', 'aMax',
                                                'xDecenter', 'yDecenter',
                                                'apertureFile'])
        return ainfo._make(apertureInfo)

    def zGetApodization(self, px, py):
        """Computes the intensity apodization of a ray from the
        apodization type and value.

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
        """Returns the graphic display aspect-ratio and the width
        (or height) of the printed page in current lens units.

        Parameters
        ----------
        filename : string
            name of the temporary file associated with the window being
            created or updated. If the temporary file is left off, then
            the default aspect-ratio and width (or height) is returned.

        Returns
        -------
        aspect : float
            aspect ratio (height/width)
        side : float
            width if ``aspect <= 1``; height if ``aspect > 1``
            (in lens units)
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
        Each window may have its own buffer data, and Zemax uses the
        filename to identify the window for which the buffer data is
        requested.

        References
        ----------
        See section "How ZEMAX calls the client" in Zemax manual [Zemax]_.

        See Also
        --------
        zSetBuffer()
        """
        cmd = "GetBuffer,{:d},{}".format(n,tempFileName)
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())
        # !!!FIX what is the proper return for this command?

    def zGetComment(self, surfNum):
        """Returns the surface comment, if any, associated with the surface

        Parameters
        ----------
        surfNum: integer
            the surface number

        Returns
        -------
        comment : string
            the comment, if any, associated with the surface
        """
        reply = self._sendDDEcommand("GetComment,{:d}".format(surfNum))
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
        The function returns ``(1,1,1)`` even if the multi-configuration
        editor is empty. This is because, the current lens in the LDE is,
        by default, set to the current configuration. The initial number of
        configurations is therefore ``1``, and the number of operators in
        the multi-configuration editor is also ``1`` (usually, ``MOFF``).

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
        return self._sendDDEcommand('GetDate').rstrip()

    def zGetExtra(self, surfNum, colNum):
        """Returns extra surface data from the Extra Data Editor

        Parameters
        ----------
        surfNum : integer
            surface number
        colNum : integer
            column number

        Returns
        -------
        value : float
            numeric data value

        See Also
        --------
        zSetExtra()
        """
        cmd="GetExtra,{sn:d},{cn:d}".format(sn=surfNum, cn=colNum)
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zGetField(self, n):
        """Returns field data for lens in Zemax DDE server

        Parameters
        ----------
        n : integer
            the field number

        Returns
        -------
        [Case: ``n=0``]

        type : integer
            0 = angles in degrees; 1 = object height; 2 = paraxial image
            height, 3 = real image height
        numFields : integer
            number of fields currently defined
        maxX : float
            values used to normalize x field coordinate
        maxY : float
            values used to normalize y field coordinate
        normMethod : integer
            normalization method (0 = radial, 1 = rectangular)

        [Case: ``0 < n <= number-of-fields``]

        xf : float
            the field-x value
        yf : float
            the field-y value
        wgt : float
            field weight value
        vdx : float
            decenter-x vignetting factor
        vdy : float
            decenter-y vignetting factor
        vcx : float
            compression-x vignetting factor
        vcy : float
            compression-y vignetting factor
        van : float
            angle vignetting factor

        Notes
        -----
        The returned tuple's content and structure is exactly same as that
        returned by ``zSetField()``

        See Also
        --------
        zSetField()
        """
        if n: # n > 0
            fd = _co.namedtuple('fieldData', ['xf', 'yf', 'wgt',
                                              'vdx', 'vdy',
                                              'vcx', 'vcy', 'van'])
        else: # n = 0
            fd = _co.namedtuple('fieldData', ['type', 'numFields',
                                              'maxX', 'maxY', 'normMethod'])
        reply = self._sendDDEcommand('GetField,'+ str(n))
        rs = reply.split(',')
        if n: # n > 0
            fieldData = fd._make([float(elem) for elem in rs])
        else: # n = 0
            fieldData = fd._make([int(elem) if (i==0 or i==1 or i==4)
                                 else float(elem) for i, elem in enumerate(rs)])
        return fieldData

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
        EFL : float
            Effective Focal Length (EFL) in lens units,
        paraWorkFNum : float
            paraxial working F/#,
        realWorkFNum : float
            real working F/#,
        paraImgHeight : float
            paraxial image height, and
        paramag : float
            paraxial magnification. See Notes.

        Notes
        ----- 
        The value of the magnification returned by this function is the
        paraxial magnification. This value doesn't depend on the acutal 
        image height or the actual location of the image surface from the 
        principal planes; instead it depends on the paraxial image height.
        For real magnification see `zGetMagnification()`.  

        See Also
        --------
        zGetSystem() :
            Use ``zGetSystem()`` to get general system data,
        zGetSystemProperty()
        ipzGetFirst()
        zGetMagnification()
        """
        fd = _co.namedtuple('firstOrderData',
                            ['EFL', 'paraWorkFNum', 'realWorkFNum',
                             'paraImgHeight', 'paraMag'])
        reply = self._sendDDEcommand('GetFirst')
        rs = reply.split(',')
        return fd._make([float(elem) for elem in rs])

    def zGetGlass(self, surfNum):
        """Returns glass data of a surface.

        Parameters
        ----------
        surfNum : integer
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
        reply = self._sendDDEcommand("GetGlass,{:d}".format(surfNum))
        rs = reply.split(',')
        if len(rs) > 1:
            glassInfo = gd._make([str(rs[i]) if i == 0 else float(rs[i])
                                                      for i in range(len(rs))])
        else:
            glassInfo = None
        return glassInfo

    def zGetGlobalMatrix(self, surfNum):
        """Returns the the matrix required to convert any local coordinates
        (such as from a ray trace) into global coordinates.

        Parameters
        ----------
        surfNum : integer
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
        Reference Surface" in the Zemax manual [Zemax]_.
        """
        gmd = _co.namedtuple('globalMatrix', ['R11', 'R12', 'R13',
                                              'R21', 'R22', 'R23',
                                              'R31', 'R32', 'R33',
                                              'Xo' ,  'Yo', 'Zo'])
        cmd = "GetGlobalMatrix,{:d}".format(surfNum)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        globalMatrix = gmd._make([float(elem) for elem in rs.split(',')])
        return globalMatrix

    def zGetIndex(self, surfNum):
        """Returns the index of refraction data for the specified surface

        Parameters
        ----------
        surfNum : integer
            surface number

        Returns
        -------
        indexData : tuple of real values
            the ``indexData`` is a tuple of index of refraction values
            defined for each wavelength in the format (n1, n2, n3, ...).
            If the specified surface is not valid, or is gradient index,
            the returned string is empty.

        See Also
        --------
        zGetIndexPrimWave()
        """
        reply = self._sendDDEcommand("GetIndex,{:d}".format(surfNum))
        rs = reply.split(",")
        indexData = [float(rs[i]) for i in range(len(rs))]
        return tuple(indexData)

    def zGetLabel(self, surfNum):
        """Returns the integer label associated with the specified surface.

        Parameters
        ----------
        surfNum : integer
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
        reply = self._sendDDEcommand("GetLabel,{:d}".format(surfNum))
        return int(float(reply.rstrip()))

    def zGetMetaFile(self, metaFileName, analysisType, settingsFile=None,
                     flag=0, timeout=None):
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
        timeout : integer, optional
            timeout in seconds (default=None, i.e. default timeout value)

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
                reply = self._sendDDEcommand(cmd, timeout)
                if 'OK' in reply.split():
                    retVal = 0
        else:
            print("Invalid analysis code '{}' passed to zGetMetaFile."
                  .format(analysisType))
        return retVal

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
        multiConData : tuple or None
            if the MCE is empty `None` is returned. Else, the exact elements of 
            ``multiConData`` depends on the value of ``config``

            If ``config > 0``
                then the elements of ``multiConData`` are:
                (value, numConfig, numRow, status, pickupRow, pickupConfig,
                scale, offset)

                The ``status`` is 0 for fixed, 1 for variable, 2 for pickup,
                & 3 for thermal pickup. If ``status`` is 2 or 3, the pickuprow &
                pickupconfig values indicate the source data for the pickup solve.

            If ``config = 0``
                then the elements of ``multiConData`` are:
                (operandType, num1, num2, num3)
                
                `num1` could be "Surface#", "Surface", "Field#", "Wave#', or 
                "Ignored". 
                `num2` could be "Object", "Extra Data Number", or "Parameter".
                `num3` could be "Property", or "Face#". 
                See [MCO]_

        References
        ----------
        .. [MCO] "Summary of Multi-Configuration Operands," Zemax manual.

        See Also
        --------
        zSetMulticon(), zGetConfig()
        """
        cmd = "GetMulticon,{config:d},{row:d}".format(config=config,row=row)
        reply = self._sendDDEcommand(cmd)
        if config: # if config > 0
            mcd = _co.namedtuple('MCD', ['value', 'numConfig', 'numRow', 'status', 
                                         'pickupRow', 'pickupConfig', 'scale',
                                         'offset'])
            rs = reply.split(",")

            if len(rs) < 8:
                if (self.zGetConfig() == (1, 1, 1)): # probably nothing set in MCE
                    return None 
                else:
                    assert False, "Unexpected reply () from Zemax.".format(reply)
            else:
                multiConData = [float(rs[i]) if (i==0 or i==6 or i==7) else int(rs[i])
                                                              for i in range(len(rs))]
        else: # if config == 0
            mcd = _co.namedtuple('MCD', ['operandType', 'num1', 'num2', 'num3'])
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0, rs[0])
        return mcd._make(multiConData)

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

    def zGetNSCData(self, surfNum, code):
        """Returns the data for NSC groups

        Parameters
        ----------
        surfNum : integer
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
        cmd = "GetNSCData,{:d},{:d}".format(surfNum,code)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscData = -1
        else:
            nscData = int(float(rs))
            if nscData == 1:
                nscObjType = self.zGetNSCObjectData(surfNum,1,0)
                if nscObjType == 'NSC_NULL': # the NSC editor is actually empty
                    nscData = 0
        return nscData

    def zGetNSCMatrix(self, surfNum, objNum):
        """Returns a tuple containing the rotation and position matrices
        relative to the NSC surface origin.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group; Use 1  for pure NSC mode
        objNum : integer
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
        cmd = "GetNSCMatrix,{:d},{:d}".format(surfNum,objNum)
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscMatrix = -1
        else:
            nscMatrix = nscmat._make([float(elem) for elem in rs.split(',')])
        return nscMatrix

    def zGetNSCObjectData(self, surfNum, objNum, code):
        """Returns the various data for NSC objects.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
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
        str_codes = (0, 1, 4)
        int_codes = (2, 3, 5, 6, 29, 101, 102, 110, 111)
        cmd = ("GetNSCObjectData,{:d},{:d},{:d}"
              .format(surfNum, objNum, code))
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

    def zGetNSCObjectFaceData(self, surfNum, objNum, faceNum, code):
        """Returns the various data for NSC object faces.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        faceNum : integer
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

            ---------------------------------------------------------------
            code  -  set/get by zGetNSCObjectFaceData/zGetNSCObjectFaceData
            ---------------------------------------------------------------
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
              .format(surfNum, objNum, faceNum, code))
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

    def zGetNSCParameter(self, surfNum, objNum, paramNum):
        """Returns NSC object's parameter data

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        paramNum : integer
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
              .format(surfNum, objNum, paramNum))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zGetNSCPosition(self, surfNum, objNum):
        """Returns position data for NSC object

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number

        Returns
        -------
        nscPos : 7-tuple (x, y, z, tilt-x, tilt-y, tilt-z, material)

        Examples
        --------
        >>> ln.zGetNSCPosition(surfNum=1, objNum=4)
        NSCPosition(x=0.0, y=0.0, z=10.0, tiltX=0.0, tiltY=0.0, tiltZ=0.0, material='N-BK7')

        See Also
        --------
        zSetNSCPosition()
        """
        nscpd = _co.namedtuple('NSCPosition', ['x', 'y', 'z',
                                               'tiltX', 'tiltY', 'tiltZ',
                                               'material'])
        cmd = ("GetNSCPosition,{:d},{:d}".format(surfNum,objNum))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPos = -1
        else:
            nscPos = nscpd._make([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                      for i in range(len(rs))])
        return nscPos

    def zGetNSCProperty(self, surfNum, objNum, faceNum, code):
        """Returns a numeric or string value from the property pages of
        objects defined in NSC editor. It mimics the ZPL function NPRO.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        faceNum : integer
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

            ---------------------------------------------------------------
            code - Datum set/get by zSetNSCProperty()/zGetNSCProperty()
            ---------------------------------------------------------------
            The following codes sets/get values to/from the NSC Editor.
              1 - Object comment (string).
              2 - Reference object number (integer).
              3 - "Inside of" object number (integer).
              4 - Object material (string).

            The following codes set/get values to/from the "Type tab" of
            the Object Properties dialog.
              0 - Object type. e.g., "NSC_SLEN" for standard lens (string).
             13 - User Defined Aperture (1 = checked, 0 = unchecked)
             14 - User Defined Aperture file name (string).
             15 - "Use Global XYZ Rotation Order" checkbox; (1 = checked,
                  0 = unchecked)
             16 - "Rays Ignore Object" checkbox; (1=checked, 0=un-checked)
             17 - "Object Is Detector" checkbox; (1=checked, 0=un-checked)
             18 - "Consider Objects" list. Argument is a string listing the
                  object numbers delimited by spaces.e.g.,"2 5 14" (string)
             19 - "Ignore Objects" list. Argument is a string listing the
                  object numbers delimited by spaces.e.g.,"1 3 7" (string)
             20 - "Use Pixel Interpolation" checkbox, (1=checked, 0=un-
                  checked).

            The following codes set/get values to/from the "Coat/Scatter
            tab" of the Object Properties dialog.
              5 - Coating name for the specified face (string)
              6 - Profile name for the specified face (string)
              7 - Scatter mode for the specified face, (0 = none,
                  1 = Lambertian, 2 = Gaussian, 3 = ABg, 4 = User Defined.)
              8 - Scatter fraction for the specified face (float)
              9 - Number of scatter rays for the specified face (integer)
             10 - Gaussian sigma for the specified face (float)
             11 - Reflect ABg data name for the specified face (string)
             12 - Transmit ABg data name for the specified face (string)
             27 - Name of the user defined scattering DLL (string)
             28 - Name of the user defined scattering data file (string)
            21-26 - Parameter values on user defined scattering DLL (float)
             29 - "Face Is" property for the specified face
                  (0 = "Object Default", 1 = "Reflective", 2 = "Absorbing")

            The following codes set/get values to/from the "Bulk Scattering
            tab" of the Object Properties dialog.
             81 - "Model" value on the bulk scattering tab (0 = "No Bulk
                  Scattering", 1 = "Angle Scattering", 2 = "DLL Defined Scattering")
             82 - Mean free path to use for bulk scattering.
             83 - Angle to use for bulk scattering.
             84 - Name of the DLL to use for bulk scattering.
             85 - Parameter value to pass to the DLL, where the face value
                  is used to specify which parameter is being defined. The
                  first parameter is 1, the second is 2, etc. (float)
             86 - Wavelength shift string (string).

            The following codes set/get values from the Diffraction tab of
            the Object Properties dialog.
             91 - "Split" value on diffraction tab (0="Don't Split By Order",
                  1="Split By Table Below", 2="Split By DLL Function")
             92 - Name of the DLL to use for diffraction splitting (string)
             93 - Start Order value (float)
             94 - Stop Order value (float)
             95 - Parameter values ondiffraction tab. These parameters are
                  passed to the diffraction splitting DLL as well as the
                  order efficiency values used by "split by table below"
                  option. The face value is used to specify which parameter
                  is being defined. The first parameter is 1, the second is
                  2, etc. (float)

            The following codes set/get values to/from the "Sources tab" of
            the Object Properties dialog.
            101 - Source object random polarization (1=checked, 0=unchecked)
            102 - Source object reverse rays option (1=checked, 0=unchecked)
            103 - Source object Jones X value
            104 - Source object Jones Y value
            105 - Source object Phase X value
            106 - Source object Phase Y value
            107 - Source object initial phase in degrees value
            108 - Source object coherence length value
            109 - Source object pre-propagation value
            110 - Source object sampling method; (0=random, 1=Sobol sampling)
            111 - Source object bulk scatter method; (0=many,1=once, 2=never)
            112 - Array mode; (0 = none, 1 = rectangular, 2 = circular,
                  3 = hexapolar, 4 = hexagonal)
            113 - Source color mode. For a complete list of the available
                  modes, see "Defining the color and spectral content of
                  sources" in the Zemax manual. The source color modes are
                  numbered starting with 0 for the System Wavelengths, and
                  then from 1 through the last model listed in the dialog
                  box control (integer)
            114-116 - Number of spectrum steps, start wavelength, and end
                      wavelength, respectively (float).
            117 - Name of the spectrum file (string).
            161-162 - Array mode integer arguments 1 and 2.
            165-166 - Array mode double precision arguments 1 and 2.
            181-183 - Source color mode arguments, for example, the XYZ
                      values of the Tristimulus (float).

            The following codes set/get values to/from the "Grin tab" of
            the Object Properties dialog.
            121 - "Use DLL Defined Grin Media" checkbox (1 = checked, 0 =
                  unchecked)
            122 - Maximum step size value (float)
            123 - DLL name (string)
            124 - Grin DLL parameters. These are the parameters passed to
                  the DLL. The face value is used to specify the parameter
                  that is being defined. The first parameter is 1, the
                  second is 2, etc (float)

            The following codes set/get values to/from the "Draw tab" of
            the Object Properties dialog.
            141 - Do not draw object checkbox (1 = checked, 0 = unchecked)
            142 - Object opacity (0 = 100%, 1 = 90%, 2 = 80%, etc.)

            The following codes set/get values to/from the "Scatter To tab"
            of the Object Properties dialog.
            151 - Scatter to method (0 = scatter to list, 1 = importance
                  sampling)
            152 - Importance Sampling target data. The argument is a string
                  listing the ray number, the object number, the size, and
                  the limit value, separated by spaces. e.g., to set the
                  Importance Sampling data for ray 3, object 6, size 3.5,
                  and limit 0.6, the string argument is "3 6 3.5 0.6".
            153 - "Scatter To List" values. Argument is a string listing
                  the object numbers to scatter to delimited by spaces,
                  such as "4 6 19" (string)

            The following codes set/get values to/from the "Birefringence
            tab" of the Object Properties dialog.
            171 - Birefringent Media checkbox (0 = unchecked, 1 = checked)
            172 - Birefringent Media Mode (0 = Trace ordinary and
                  extraordinary rays, 1 = Trace only ordinary rays, 2 =
                  Trace only extraordinary rays, and 3 = Waveplate mode)
            173 - Birefringent Media Reflections status (0 = Trace
                  reflected and refracted rays, 1 = Trace only refracted
                  rays, and 2 = Trace only reflected rays)
            174-176 - Ax, Ay, and Az values (float)
            177 - Axis Length (float)
            200 - Index of refraction of an object (float)
            201-203 - nd (201), vd (202), and dpgf (203) parameters of an
                      object using a model glass.

            end-of-table

        See Also
        --------
        zSetNSCProperty()
        """
        cmd = ("GetNSCProperty,{:d},{:d},{:d},{:d}"
                .format(surfNum, objNum, code, faceNum))
        reply = self._sendDDEcommand(cmd)
        nscPropData = _process_get_set_NSCProperty(code, reply)
        return nscPropData

    def zGetNSCSettings(self):
        """Returns the maximum number of intersections, segments, nesting
        level, minimum absolute intensity, minimum relative intensity, glue
        distance, miss ray distance, ignore errors flag used for NSC ray
        tracing.

        Parameters
        ----------
        None

        Returns
        -------
        maxIntersec : integer
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
        nscSetData = _co.namedtuple('nscSettings', ['maxIntersec', 'maxSeg', 'maxNest',
                                                    'minAbsI', 'minRelI', 'glueDist',
                                                    'missRayLen', 'ignoreErr'])
        return nscSetData._make(nscSettingsData)

    def zGetNSCSolve(self, surfNum, objNum, param):
        """Returns the current solve status and settings for NSC position
        and parameter data

        Parameters
        ----------
        surfNum : integer
            surface number of NSC group; use 1 if program mode is pure NSC
        objNum : integer
            object number
        param : integer
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

            The status value is 0 for fixed, 1 for variable, and 2 for a
            pickup solve.

            Only when the staus is a pickup solve is the other data
            meaningful.

            -1 if it a BAD COMMAND

        See Also
        --------
        zSetNSCSolve()
        """
        nscSolveData = -1
        cmd = ("GetNSCSolve,{:d},{:d},{:d}"
               .format(surfNum, objNum, param))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if 'BAD COMMAND' not in rs:
            nscSolveData = tuple([float(e) if i in (3,4) else int(float(e))
                                 for i,e in enumerate(rs.split(","))])
        return nscSolveData

    def zGetOperand(self, row, column):
        """Return the operand data from the Merit Function Editor

        Parameters
        ----------
        row : integer
            operand row number in the MFE
        column : integer
            column number

        Returns
        -------
        operandData : integer/float/string
            opernadData's type depends on ``column`` argument if
            successful, else -1.

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
        zGetOperandRow():
            Returns all values from a row in MFE
        zOperandValue():
            Returns the value of any optimization operand, even if the
            operand is not currently in the merit function.
            Use ``zOperandValue()`` instead of ``zGetOperand()`` if you
            just want to observe/retrieve (instead of optimizing) any
            operand variable.
        zOptimize() :
            To update merit function prior to calling ``zGetOperand()``,
            call ``zOptimize()`` with the number of cycles set to -1
        ipzGetMFE() :
            prints/ returns the MFE parameter suitable for interactive
            environment
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
            if ``nlsPol > 0``, then default pol. state is unpolarized
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
        """Trace a single polarized ray defined by the normalized field
        height, pupil height, electric field magnitude and relative phase.

        If ``Ex``, ``Ey``, ``Phax``, ``Phay`` are all zero, two orthogonal
        rays are traced; the resulting transmitted intensity is averaged.

        Parameters
        ----------
        waveNum : integer
            wavelength number as in the wavelength data editor
        mode : integer (0/1)
            0 = real, 1 = paraxial
        surf : integer
            surface to trace the ray to. if -1, surf is the image plane
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
        error : integer
            0, if the ray traced successfully;
            +ve number indicates ray missed the surface
            -ve number indicates ray total internal reflected (TIR)
            at the surface given by the absolute value of the ``error``
        intensity : float
            the transmitted intensity of the ray, normalized to an input
            electric field intensity of unity. The transmitted intensity
            accounts for surface, thin film, and bulk absorption effects,
            but does not consider whether or not the ray was vignetted.
        Exr,Eyr,Ezr : float
            real parts of the electric field components
        Exi,Eyi,Ezi : float
            imaginary parts of electric field components


        For unploarized rays, only the ``error`` and ``intensity`` are
        relevant.

        Examples
        --------
        To trace the real unpolarized marginal ray to the image surface at
        wavelength 2, the function would be:

        >>> ln.zGetPolTrace(2, 0, -1, 0.0, 0.0, 0.0, 1.0, 0, 0, 0, 0)

        .. _notes-GetPolTrace:

        Notes
        -----
        1. The quantity ``Ex*Ex + Ey*Ey`` should have a value of 1.0
           although any values are accepted.
        2. There is an important exception to the above rule -- If ``Ex``,
           ``Ey``, ``Phax``, ``Phay`` are all zero, Zemax will trace two
           orthogonal rays, and the resulting transmitted intensity
           will be averaged.
        3. Always check to verify the ray data is valid (check ``error``)
           before using the rest of the data in the tuple.
        4. Use of ``zGetPolTrace()`` has significant overhead as only one
           ray per DDE call is traced. Please refer to the Zemax manual for
           more details.

        See Also
        --------
        zGetPolTraceDirect(), zGetTrace(), zGetTraceDirect()
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f},".format(px=px,py=py)
        args4 = "{Ex:1.4f},{Ey:1.4f},".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = "GetPolTrace," + args1 + args2 + args3 + args4 + args5
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        polRayTraceData = [int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)]
        rtd = _co.namedtuple('polRayTraceData', ['error', 'intensity',
                                                 'Exr', 'Eyr', 'Ezr',
                                                 'Exi', 'Eyi', 'Ezi'])
        polRayTraceData = rtd._make(polRayTraceData)
        return polRayTraceData

    def zGetPolTraceDirect(self, waveNum, mode, startSurf, stopSurf,
                           x, y, z, l, m, n, Ex, Ey, Phax, Phay):
        """Trace a single polarized ray defined by the ``x``, ``y``,
        ``z``, ``l``, ``m`` and ``n`` coordinates on any starting
        surface as well as electric field magnitude and relative phase.

        If ``Ex``, ``Ey``, ``Phax``, ``Phay`` are all zero, Zemax will
        trace two orthogonal rays and the resulting transmitted intensity
        will be averaged.

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
        error : integer
            0, if the ray traced successfully;
            +ve number indicates ray missed the surface
            -ve number indicates ray total internal reflected (TIR)
            at the surface given by the absolute value of the ``error``
        intensity : float
            the transmitted intensity of the ray, normalized to an input
            electric field intensity of unity. The transmitted intensity
            accounts for surface, thin film, and bulk absorption effects,
            but does not consider whether or not the ray was vignetted.
        Exr,Eyr,Ezr : float
            real parts of the electric field components
        Exi,Eyi,Ezi : float
            imaginary parts of electric field components


        For unploarized rays, only the ``error`` and ``intensity`` are
        relevant.

        Notes
        -----
        Refer to the notes (notes-GetPolTrace_) of ``zGetPolTrace()``

        See Also
        --------
        zGetPolTraceDirect(), zGetTrace(), zGetTraceDirect()
        """
        args0 = "{wN:d},{m:d},".format(wN=waveNum,m=mode)
        args1 = "{sa:d},{sd:d},".format(sa=startSurf,sd=stopSurf)
        args2 = "{x:1.20g},{y:1.20g},{z:1.20g},".format(x=x,y=y,z=z)
        args3 = "{l:1.20g},{m:1.20g},{n:1.20g},".format(l=l,m=m,n=n)
        args4 = "{Ex:1.4f},{Ey:1.4f},".format(Ex=Ex,Ey=Ey)
        args5 = "{Phax:1.4f},{Phay:1.4f}".format(Phax=Phax,Phay=Phay)
        cmd = ("GetPolTraceDirect," + args0 + args1 + args2 + args3
                                    + args4 + args5)
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        polRayTraceData = [int(elem) if i==0 else float(elem)
                                   for i,elem in enumerate(rs)]
        rtd = _co.namedtuple('polRayTraceData', ['error', 'intensity',
                                                 'Exr', 'Eyr', 'Ezr',
                                                 'Exi', 'Eyi', 'Ezi'])
        polRayTraceData = rtd._make(polRayTraceData)
        return polRayTraceData

    def zGetPupil(self):
        """Return the pupil data such as aperture type, ENPD, EXPD, etc.

        Parameters
        ----------
        None

        Returns
        -------
        aType : integer
            the system aperture defined as follows:

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
            entrance pupil position from the first surface (in lens units)
        EXPD : float
            exit pupil diameter (in lens units)
        EXPP : float
            exit pupil position from the image plane (in lens units)
        apodization_type : integer
            the apodization type is indicated as follows:

                * 0 = none
                * 1 = Gaussian
                * 2 = Tangential/Cosine cubed

        apodization_factor : float
            number shown on general data dialog box
        """
        pupild = _co.namedtuple('PupilData', ['aType', 'value', 'ENPD',
                                              'ENPP', 'EXPD', 'EXPP',
                                              'apoType', 'apoFactor'])
        reply = self._sendDDEcommand('GetPupil')
        rs = reply.split(',')
        pupilData = pupild._make([int(elem) if (i==0 or i==6)
                                 else float(elem) for i, elem in enumerate(rs)])
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
            -1 if Zemax could not copy the lens data LDE to the server;
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

    def zGetSag(self, surfNum, x, y):
        """Return the sag of the surface at coordinates (x,y) in lens units

        Parameters
        ----------
        surfNum : integer
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
        cmd = "GetSag,{:d},{:1.20g},{:1.20g}".format(surfNum,x,y)
        reply = self._sendDDEcommand(cmd)
        sagData = reply.rsplit(",")
        return (float(sagData[0]),float(sagData[1]))

    def zGetSequence(self):
        """Returns the sequence numbers of the lens in the server and in
        the LDE

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
            the name of the output file passed by Zemax to the client.
            Zemax uses this name to identify for the window for which the
            ``zGetSettingsData()`` request is for.
        number : integer
            the data number used by the previous ``zSetSettingsData()``
            call. Currently, only ``number=0`` is supported.

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

    def zGetSolve(self, surfNum, code):
        """Returns data about solves and/or pickups on the surface

        Parameters
        ----------
        surfNum : integer
            surface number
        code : integer
            indicating the surface parameter, such as curvature, thickness,
            glass, conic, semi-diameter, etc. (Refer to the table
            surf_param_codes_for_getsolve_ or use the surface
            parameter mnemonic codes with signature `ln.SOLVE_SPAR_XXX`, e.g. 
            `ln.SOLVE_SPAR_CURV`, `ln.SOLVE_SPAR_THICK`, etc. The `SPAR` 
            stands for surface parameter.

        Returns
        -------
        solveData : tuple
            tuple is depending on the code value according to the table;
            returns -1 if error occurs

        Examples
        -------- 
        >>> solvetype, param1, param2, param3, pickup = ln.zGetSolve(3, ln.SOLVE_SPAR_THICK)

        In the above example, since the solve is on Thickness (code=ln.SOLVE_SPAR_THICK), 
        if the `solvetype` is "Position" (7), then `param1` is "From Surface", 
        `param2` is "Length", and `param3` and `pickup` are un-specified. So, a typical 
        output could be `(7, 3.0, 0.0, 0.0, 0)`. Instead of "Position", if the `solvetype`
        is "Pickup" (5), then `param1` is "From Surface", `param2` is "Scale Factor", 
        `param3` is "Offset", and `pickup` is "Pickup column"     

        Notes
        -----

        .. _surf_param_codes_for_getsolve:

        ::

            Table : Surface parameter codes for zGetsolve() and zSetSolve()

            ------------------------------------------------------------------------------
               code           - Datum set/get by zGetSolve()/zSetSolve()
            ------------------------------------------------------------------------------
            0 (curvature)     - solvetype, param1, param2, pickupcolumn
            1 (thickness)     - solvetype, param1, param2, param3, pickupcolumn
            2 (glass)         - solvetype (for solvetype = 0);
                                solvetype, Index, Abbe, Dpgf (for solvetype = 1, model glass);
                                solvetype, pickupsurf (for solvetype = 2, pickup);
                                solvetype, index_offset, abbe_offset (for solvetype = 4, offset);
                                solvetype (for solvetype=all other values)
            3 (semi-diameter) - solvetype, pickupsurf, pickupcolumn
            4 (conic)         - solvetype, pickupsurf, pickupcolumn
            5-16 (param 1-12) - solvetype, pickupsurf, offset, scalefactor, pickupcolumn
            17 (parameter 0)  - solvetype, pickupsurf, offset, scalefactor, pickupcolumn
            1001+ (extra      - solvetype, pickupsurf, scalefactor, offset, pickupcolumn
            data values 1+)     

            end-of-table

        The ``solvetype`` is an integer code, & the parameters have
        meanings that depend upon the solve type; see the chapter
        "SOLVES" in the Zemax manual for details.

        See Also
        --------
        zSetSolve(), zGetNSCSolve(), zSetNSCSolve()
        """
        cmd = "GetSolve,{:d},{:d}".format(surfNum,code)
        reply = self._sendDDEcommand(cmd)
        solveData = _process_get_set_Solve(reply)
        return solveData

    def zGetSurfaceData(self, surfNum, code, arg2=None):
        """Gets surface data on a sequential lens surface.

        Parameters
        ----------
        surfNum : integer
            the surface number
        code : integer
            integer code (see table surf_data_codes_ below). You may also 
            use the surface data mnemonic codes with signature ln.SDAT_XXX, 
            e.g. ln.SDAT_TYPE, ln.SDAT_CURV, ln.SDAT_THICK, etc 
        arg2 : integer, optional
            required for item ``codes`` above 70.

        Returns
        -------
        surface_data : string or numeric
            the returned data depends on the ``code``. Refer to the table
            surf_data_codes_ for details.

        Notes
        -----

        .. _surf_data_codes:

        ::

            Table : Surface data codes for getter and setter of SurfaceData

            ---------------------------------------------------------------
            Code   -  Datum set/get by zSetSurfaceData()/zGetSurfaceData()
            ---------------------------------------------------------------
            0      - Surface type name (string)
            1      - Comment (string)
            2      - Curvature (numeric)
            3      - Thickness (numeric)
            4      - Glass (string)
            5      - Semi-Diameter (numeric)
            6      - Conic (numeric)
            7      - Coating (string)
            8      - Thermal Coefficient of Expansion (TCE)
            9      - User-defined .dll (string)
            20     - Ignore surface flag. 0 for not ignored; 1 for ignored
            51     - Before tilt and decenter order; 0 for Decenter
                     then Tilt; 1 for Tilt then Decenter
            52     - Before decenter x
            53     - Before decenter y
            54     - Before tilt x
            55     - Before tilt y
            56     - Before tilt z 
            60     - After status. 0 for explicit; 1 for pickup current surface; 
                     2 for reverse current surface; 3 for pickup previous surface; 
                     4 for reverse previous surface, etc.
            61     - After tilt and decenter order; 0 for Decenter
                     then Tilt, 1 for Tilt then Decenter
            62     - After decenter x
            63     - After decenter y 
            64     - After tilt x
            65     - After tilt y 
            66     - After tilt z
            70     - Use Layer Multipliers and Index Offsets. Use 1 for
                     true, 0 for false.
            71     - Layer Multiplier value. The coating layer number is
                     defined by ``arg2``
            72     - Layer Multiplier status. Use 0 for fixed; 1 for
                     variable; or n+1 for pickup from layer n. The coating
                     layer number is defined by ``arg2``
            73     - Layer Index Offset value. The coating layer number is
                     defined by ``arg2``
            74     - Layer Index Offset status. Use 0 for fixed; 1 for
                     variable, or n+1 for pickup from layer n. The coating
                     layer number is defined by ``arg2``
            75     - Layer Extinction Offset value. The coating layer
                     number is defined by ``arg2``
            76     - Layer Extinction Offset status. Use 0 for fixed; 1 for
                     variable, or n+1 for pickup from layer n. The coating
                     layer number is defined by ``arg2``
            Other  - Reserved for future expansion of this feature.

            end-of-table

        See Also
        --------
        zSetSurfaceData(), zGetSurfaceParameter()
        """
        if arg2 is None:
            cmd = "GetSurfaceData,{sN:d},{c:d}".format(sN=surfNum,c=code)
        else:
            cmd = "GetSurfaceData,{sN:d},{c:d},{a:d}".format(sN=surfNum,
                                                                 c=code,a=arg2)
        reply = self._sendDDEcommand(cmd)
        if code in (0,1,4,7,9):
            surfaceDatum = reply.rstrip()
        else:
            surfaceDatum = float(reply)
        return surfaceDatum

    def zGetSurfaceDLL(self, surfNum):
        """Return the name of the DLL if the surface is a user defined type

        Parameters
        ----------
        surfNum : integer
            surface number of the user defined surface

        Returns
        -------
        dllName : string
            The name of the defining DLL
        surfaceName : string
            surface name displayed by the DLL in the surface type column of
            the LDE
        """
        cmd = "GetSurfaceDLL,{sN:d}".format(surfNum)
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        return (rs[0],rs[1])

    def zGetSurfaceParameter(self, surfNum, param):
        """Return the surface parameter data for the surface associated
        with the given surface number `surfNum`

        Parameters
        ----------
        surfNum : integer
            surface number of the surface
        param : integer
            parameter number ('Par' in LDE) being queried

        Returns
        --------
        paramData : float
            the parameter value

        See Also
        --------
        zGetSurfaceData() :
             To get thickness, radius, glass, semi-diameter, conic, etc,
        zSetSurfaceParameter()
        """
        cmd = ("GetSurfaceParameter,{sN:d},{p:d}"
               .format(sN=surfNum,p=param))
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zGetSystem(self):
        """Returns a number of general system data (General Lens Data)

        Parameters
        ----------
        None

        Returns
        -------
        numSurfs : integer
            number of surfaces
        unitCode : integer
            lens units code (0, 1, 2, or 3 for mm, cm, in, or M)
        stopSurf : integer
            the stop surface number
        nonAxialFlag : integer
            flag to indicate if system is non-axial symmetric (0 for axial,
            1 if not axial);
        rayAimingType : integer
            ray aiming type (0, 1, or 2 for off, paraxial or real)
        adjustIndex : integer
            adjust index data to environment flag (0 if false, 1 if true)
        temp : float
            the current temperature
        pressure : float
            the current pressure
        globalRefSurf : integer
            the global coordinate reference surface number

        Notes
        -----
        The returned data structure is same as the data structure returned
        by the ``zSetSystem()`` method

        See Also
        --------
        zGetFirst() :
            to get first order lens data such as EFL, F/#, etc.
        zSetSystem(), zGetSystemProperty(), zGetSystemAper(),
        zGetAperture(), zSetAperture()
        """
        sdt = _co.namedtuple('systemData' , ['numSurf', 'unitCode',
                                             'stopSurf', 'nonAxialFlag',
                                             'rayAimingType', 'adjustIndex',
                                             'temp', 'pressure',
                                             'globalRefSurf'])
        reply = self._sendDDEcommand("GetSystem")
        rs = reply.split(',')
        systemData = sdt._make([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zGetSystemAper(self):
        """Gets system aperture data -- aperture type, stopSurf and value.

        Returns
        -------
        aType : integer
            indicating the system aperture as follows:

            | 0 = entrance pupil diameter (EPD)
            | 1 = image space F/#         (IF/#)
            | 2 = object space NA         (ONA)
            | 3 = float by stop           (FBS)
            | 4 = paraxial working F/#    (PWF/#)
            | 5 = object cone angle       (OCA)

        stopSurf : integer
            stop surface
        value : float
            if aperture type is "float by stop" value is stop surface
            semi-diameter else value is the sytem aperture

        Notes
        -----
        The returned tuple is the same as the returned tuple of
        ``zSetSystemAper()``

        See Also
        --------
        zGetSystem(), zSetSystemAper()
        """
        sad = _co.namedtuple('systemAper', ['apertureType', 'stopSurf', 'value'])
        reply = self._sendDDEcommand("GetSystemAper")
        rs = reply.split(',')
        systemAperData = sad._make([float(elem) if i==2 else int(float(elem))
                                    for i, elem in enumerate(rs)])
        return systemAperData

    def zGetSystemProperty(self, code):
        """Returns properties of the system, such as system aperture, field,
        wavelength, and other data, based on the integer ``code`` passed.

        This function mimics the ZPL function ``SYPR``.

        Parameters
        ----------
        code : integer
            value that defines the specific system property. (see the table
            system_property_codes_ below).

        Returns
        -------
        sysPropData : string or numeric
            Returned system property data

        Notes
        -----

        .. _system_property_codes

        ::

            Table : System property codes

            ---------------------------------------------------------------
            Code    set/get by zSetSystemProperty()/zGetSystemProperty()
            ---------------------------------------------------------------
              4   - Adjust Index Data To Environment (0:off, 1:on)
             10   - Aperture Type code. (0:EPD, 1:IF/#, 2:ONA, 3:FBS,
                    4:PWF/#, 5:OCA)
             11   - Aperture Value (stop surface semi-diameter if aperture
                    type is FBS, else system aperture)
             12   - Apodization Type code. (0:uniform, 1:Gaussian, 2:cosine
                    cubed)
             13   - Apodization Factor
             14   - Telecentric Object Space (0:off, 1:on)
             15   - Iterate Solves When Updating (0:off, 1:on)
             16   - Lens Title
             17   - Lens Notes
             18   - Afocal Image Space (0:off or "focal mode", 1:on or
                    "afocal mode")
             21   - Global coordinate reference surface
             23   - Glass catalog list (Use a string or string variable
                    with the glass catalog name, such as "SCHOTT". To
                    specify multiple catalogs use a single string or string
                    variable containing names separated by spaces, such as
                    "SCHOTT HOYA OHARA".)
             24   - System Temperature in degrees Celsius.
             25   - System Pressure in atmospheres.
             26   - Reference OPD method. (0:absolute, 1:infinity, 2:exit
                    pupil, 3:absolute 2.)
             30   - Lens Unit code (0:mm, 1:cm, 2:inches, 3:Meters)
             31   - Source Unit Prefix (0:Femto, 1:Pico, 2:Nano, 3:Micro,
                    4:Milli, 5:None, 6:Kilo, 7:Mega, 8:Giga, 9:Tera)
             32   - Source Units. (0:Watts, 1:Lumens, 2:Joules)
             33   - Analysis Unit Prefix (0:Femto, 1:Pico, 2:Nano, 3:Micro,
                    4:Milli, 5:None, 6:Kilo, 7:Mega, 8:Giga, 9:Tera)
             34   - Analysis Units "per" Area. (0:mm^2, 1:cm^2, 2:inches^2,
                    3:Meters^2, 4:feet^2)
             35   - MTF Units code. (0:cycles per millimeter, 1:cycles per
                    milliradian.
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
             64   - Convert thin film phase to ray equivalent (0:no, 1:yes)
             65   - Unpolarized. (0:no, 1:yes)
             66   - Method. (0:X-axis, 1:Y-axis, 2:Z-axis)
             70   - Ray Aiming. (0:off, 1:on (paraxial), 2:aberrated (real))
             71   - Ray aiming pupil shift x.
             72   - Ray aiming pupil shift y.
             73   - Ray aiming pupil shift z.
             74   - Use Ray Aiming Cache. (0:no, 1:yes)
             75   - Robust Ray Aiming. (0:no, 1:yes)
             76   - Scale Pupil Shift Factors By Field. (0:no, 1:yes)
             77   - Ray aiming pupil compress x.
             78   - Ray aiming pupil compress y.
             100  - Field type code. (0=angl, 1=obj ht, 2=parx img ht,
                    3=rel img ht)
             101  - Number of fields.
             102,103 - The field number is value1, value2 is the field x,
                    y coordinate
             104  - The field number is value1, value2 is the field weight
             105,106 - The field number is value1, value2 is the field
                    vignetting decenter x, decenter y
             107,108 - The field number is value1, value2 is the field
                    vignetting compression x, compression y
             109  - The field number is value1, value2 is the field
                    vignetting angle
             110  - The field normalization method, value 1 is 0 for radial
                    and 1 for rectangular
             200  - Primary wavelength number.
             201  - Number of wavelengths
             202  - The wavelength number is value1, value 2 is the
                    wavelength in micrometers.
             203  - The wavelength number is value1, value 2 is the
                    wavelength weight
             901  - The number of CPU's to use in multi-threaded
                    computations, such as optimization. (0=default). See
                    the manual for details.

            NOTE: Currently Zemax returns just "0" for the codes: 102,103,
                  104,105, 106,107,108,109, and 110. This is unexpected!
                  So, PyZDDE will return the reply (string) as-is for the
                  user to handle.

            end-of-table

        See Also
        --------
        zSetSystemProperty(), zGetFirst()
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
            name of the file to be created including the full path and
            extension.
        analysisType : string
            3 letter case-sensitive label that indicates the type of the
            analysis to be performed. They are identical to the button
            codes. If no label is provided or recognized, a standard
            raytrace will be generated
        settingsFile : string
            If ``settingsFile`` is valid, Zemax will use or save the
            settings used to compute the text file, depending upon the
            value of the flag parameter
        flag : integer (0, 1, or 2)
            0 = default settings used for the text;
            1 = settings provided in the settings file, if valid, else
                default;
            2 = settings provided in the settings file, if valid, will be
                used and the settings box for the requested feature will
                be displayed. After the user makes any changes to the
                settings the text will then be generated using the new
                settings. Please see the manual for more details
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
        No matter what the flag value is, if a valid file name is provided
        for ``settingsFile``, the settings used will be written to the
        settings file, overwriting any data in the file.

        See Also
        --------
        zGetMetaFile(), zOpenWindow()
        """
        retVal = -1
        if settingsFile:
            settingsFile = settingsFile
        else:
            settingsFile = ''
        # Check if the file path is valid and has extension
        if _os.path.isabs(textFileName) and _os.path.splitext(textFileName)[1]!='':
            cmd = 'GetTextFile,"{tF}",{aT},"{sF}",{fl:d}'.format(tF=textFileName,
                                    aT=analysisType,sF=settingsFile,fl=flag)
            reply = self._sendDDEcommand(cmd, timeout)
            if 'OK' in reply.split():
                retVal = 0
        return retVal

    def zGetTol(self, operNum):
        """Returns the tolerance data

        Parameters
        ----------
        operNum : integer
            0 or the tolerance operand number (row number in the tolerance
            editor, when greater than 0)

        Returns
        -------
        toleranceData : single number or a 6-tuple
            It is a number or a 6-tuple, depending upon ``operNum``
            as follows:

              * if ``operNum==0``, then toleranceData = number where
                ``number`` is the number of tolerance operands defined.
              * if ``operNum > 0``, then toleranceData =
                (tolType, int1, int2, min, max, int3)

        See Also
        --------
        zSetTol(), zSetTolRow()
        """
        reply = self._sendDDEcommand("GetTol,{:d}".format(operNum))
        if operNum == 0:
            toleranceData = int(float(reply.rstrip()))
            if toleranceData == 1:
                reply = self._sendDDEcommand("GetTol,1")
                tolType = reply.rsplit(",")[0]
                if tolType == 'TOFF': # the tol editor is actually empty
                    toleranceData = 0
        else:
            toleranceData = _process_get_set_Tol(operNum,reply)
        return toleranceData

    def zGetTrace(self, waveNum, mode, surf, hx, hy, px, py):
        """Trace a ray defined by its normalized field and pupil heights
        as well as wavelength through the lens in the Zemax DDE server

        Parameters
        ----------
        waveNum : integer
            wavelength number as in the wavelength data editor
        mode : integer
            0 = real; 1 = paraxial
        surf : integer
            surface to trace the ray to. Usually, the ray data is only
            needed at the image surface; setting the surface number to
            -1 will yield data at the image surface.
        hx : real
            normalized field height along x axis
        hy : real
            normalized field height along y axis
        px : real
            normalized height in pupil coordinate along x axis
        py : real
            normalized height in pupil coordinate along y axis

        Returns
        -------
        error : integer
            0 = ray traced successfully;
            +ve number = the ray missed the surface;
            -ve number = the ray total internal reflected (TIR) at surface
                         given by the absolute value of the ``error``
        vig : integer
            the first surface where the ray was vignetted. Unless an error
            occurs at that surface or subsequent to that surface, the ray
            will continue to trace to the requested surface.
        x, y, z : reals
            coordinates of the ray on the requested surface
        l, m, n : reals
            the direction cosines after refraction into the media following
            the requested surface.
        l2, m2, n2 : reals
            the surface intercept direction normals at requested surface
        intensity : real
            the relative transmitted intensity of the ray, including any
            pupil or surface apodization defined.

        Examples
        --------
        To trace the real chief ray to surface 5 for wavelength 3:

        >>> rayTraceData = ln.zGetTrace(3,0,5,0.0,1.0,0.0,0.0)
        >>> error,vig,x,y,z,l,m,n,l2,m2,n2,intensity = rayTraceData

        Notes
        -----
        1. Always check to verify the ray data is valid ``error`` before
           using the rest of the returned parameters
        2. Use of ``zGetTrace()`` has significant overhead as only 1 ray
           per DDE call is traced. Use ``arraytrace.zGetTraceArray()`` for
           tracing large number of rays.

        See Also
        --------
        arraytrace.zGetTraceArray(), zGetTraceDirect(), zGetPolTrace(),
        zGetPolTraceDirect()
        """
        args1 = "{wN:d},{m:d},{s:d},".format(wN=waveNum,m=mode,s=surf)
        args2 = "{hx:1.4f},{hy:1.4f},".format(hx=hx,hy=hy)
        args3 = "{px:1.4f},{py:1.4f}".format(px=px,py=py)
        cmd = "GetTrace," + args1 + args2 + args3
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rayData = [int(elem) if (i==0 or i==1)
                                  else float(elem) for i,elem in enumerate(rs)]
        rtd = _co.namedtuple('rayTraceData', ['error', 'vig', 'x', 'y', 'z',
                                              'dcos_l', 'dcos_m', 'dcos_n',
                                              'dnorm_l2', 'dnorm_m2', 'dnorm_n2',
                                              'intensity'])
        rayTraceData = rtd._make(rayData)
        return rayTraceData

    def zGetTraceDirect(self, waveNum, mode, startSurf, stopSurf, x, y, z, l, m, n):
        """Trace a (single) ray defined by ``x``, ``y``, ``z``, ``l``,
        ``m`` and ``n`` coordinates on any starting surface as well as
        wavelength number, mode and the surface to.

        ``zGetTraceDirect`` provides a more direct access to the Zemax
        ray tracing engine than ``zGetTrace``.

        Parameters
        ----------
        waveNum : integer
            wavelength number as in the wavelength data editor
        mode : integer
            0 = real, 1 = paraxial
        startSurf : integer
            starting surface of the ray
        stopSurf : integer
            stopping surface of the ray
        x, y, z, : floats
            coordinates of the ray at the starting surface
        l, m, n : floats
            direction cosines to the entrance pupil aim point for the x-,
            y-, z- direction cosines respectively

        Returns
        -------
        error : integer
            0 = ray traced successfully;
            +ve number = the ray missed the surface;
            -ve number = the ray total internal reflected (TIR) at surface
                         given by the absolute value of the ``error``
        vig : integer
            the first surface where the ray was vignetted. Unless an error
            occurs at that surface or subsequent to that surface, the ray
            will continue to trace to the requested surface.
        x, y, z : reals
            coordinates of the ray on the requested surface
        l, m, n : reals
            the direction cosines after refraction into the media following
            the requested surface.
        l2, m2, n2 : reals
            the surface intercept direction normals at requested surface
        intensity : real
            the relative transmitted intensity of the ray, including any
            pupil or surface apodization defined.

        Notes
        -----
        Normally, rays are defined by the normalized field and pupil
        coordinates hx, hy, px, and py. Zemax takes these normalized
        coordinates and computes the object coordinates (x, y, and z) and
        the direction cosines to the entrance pupil aim point (l, m, and n;
        for the x-, y-, and z-direction cosines, respectively). However,
        there are times when it is more appropriate to trace rays by direct
        specification of x, y, z, l, m, and n. The direct specification has
        the added flexibility of defining the starting surface for the ray
        anywhere in the optical system.
        """
        args1 = "{wN:d},{m:d},".format(wN=waveNum,m=mode)
        args2 = "{sa:d},{sp:d},".format(sa=startSurf,sp=stopSurf)
        args3 = "{x:1.20f},{y:1.20f},{z:1.20f},".format(x=x,y=y,z=z)
        args4 = "{l:1.20f},{m:1.20f},{n:1.20f}".format(l=l,m=m,n=n)
        cmd = "GetTraceDirect," + args1 + args2 + args3 + args4
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        rtd = _co.namedtuple('rayTraceData', ['err', 'vig', 'x', 'y', 'z',
                                              'dcos_l', 'dcos_m', 'dcos_n',
                                              'dnorm_l2', 'dnorm_m2', 'dnorm_n2',
                                              'intensity'])
        rayTraceData = rtd._make([int(elem) if (i==0 or i==1)
                                  else float(elem) for i,elem in enumerate(rs)])
        return rayTraceData

    def zGetUDOSystem(self, bufferCode):
        """Load a particular lens from the optimization function memory
        into the Zemax server's memory. This will cause Zemax to retrieve
        the correct lens from system memory, and all subsequent DDE calls
        will be for actions (such as ray tracing) on this lens. The only
        time this item name should be used is when implementing a User
        Defined Operand, or UDO.

        Parameters
        ----------
        bufferCode: integer
            The ``buffercode`` is an integer value provided by Zemax to the
            client that uniquely identifies the correct lens.

        Returns
        ------
          ?

        Notes
        -----
        Once the data is computed, up to 1001 values may be sent back to
        the server, and ultimately to the optimizer within Zemax, with the
        ``zSetUDOItem()`` function.

        See Also
        --------
        zSetUDOItem()
        """
        cmd = "GetUDOSystem,{:d}".format(bufferCode)
        reply = self._sendDDEcommand(cmd)
        return _regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zGetUpdate(self):
        """Update the lens. Zemax recomputes all pupil positions, solves,
        and index data.

        Parameters
        ----------
        None

        Returns
        -------
        status :  integer
            0 = Zemax successfully updated the lens;
            -1 = No raytrace performed;
            -998 = Command timed out

        Notes
        -----
        To update the merit function, use the ``zOptimize()`` with the
        number of cycles set to -1.

        See Also
        --------
        zGetRefresh(), zOptimize(), zPushLens()
        """
        status,ret = -998, None
        ret = self._sendDDEcommand("GetUpdate")
        if ret != None:
            status = int(ret)  #Note: Zemax returns -1 if GetUpdate fails.
        return status

    def zGetVersion(self):
        """Get the version of Zemax

        Parameters
        ----------
        None

        Returns
        -------
        version : integer
            The application release date in YY-MM-DD format
        """
        return int(self._sendDDEcommand("GetVersion"))

    def zGetWave(self, n):
        """Get wavelength data

        There are 2 ways of using this function:

            * `zGetWave(0)-> waveData` Or
            * `zGetWave(wavelengthNumber)-> waveData`

        Returns
        -------
        if n==0 : 2-tuple
            * primary : (integer) number indicating the primary wavelength
            * number  : (integer) number_of_wavelengths currently defined
        if 0 < n <= number_of_wavelengths : 2-tuple
            * wavelength : (float) value of the specific wavelength (in micrometers)
            * weight : (float) weight of the specific wavelength

        Notes
        -----
        1. The returned tuple is exactly same in structure and contents to that 
           returned by ``zSetWave()``.
        2. Wavelength data are always measured in micrometers referenced to "air"  
           at the system temperature and pressure.

        See Also
        --------
        zSetWave(), zSetWaveTuple(), zGetWaveTuple(), zGetPrimaryWave()
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

    def zHammer(self, numOfCycles, algorithm, timeout=60):
        """Calls the Hammer optimizer

        Parameters
        ---------
        numOfCycles : integer
            the number of cycles to run. If ``numOfCycles < 1``, no
            optimization is performed, but Zemax updates all operands in
            the MFE and returns the merit-function
        algorithm : integer
            0 = Damped Least Squares; 1 = Orthogonal descent
        timeout : integer
            timeout in seconds (default = 60 seconds)

        Returns
        -------
        finalMeritFn : float
            the final merit function.

        Notes
        -----
        1. A returned value of 9.0E+009 indicates that the optimization
           failed. This is usually because the lens or merit function could
           not be evaluated.
        2. The number of cycles should be kept small enough to allow the
           algorithm to complete and return before the DDE communication
           times out, or an error will occur. One possible way to achieve
           high number of cycles could be to call ``zHammer()`` multiple
           times in a loop, each time comparing the returned merit function
           with few of the previously returned (& stored) merit functions
           to determine if an optimum has been attained.

        See Also
        --------
        zOptimize(),  zLoadMerit(), zsaveMerit()
        """
        cmd = "Hammer,{:1.2g},{:d}".format(numOfCycles, algorithm)
        reply = self._sendDDEcommand(cmd, timeout)
        return float(reply.rstrip())

    def zImportExtraData(self, surfNum, fileName):
        """Imports extra data and grid surface data values into an existing
        surface.

        Parameters
        ----------
        surfNum : integer
            surface number
        fileName : string
            file name (of an ASCII file)

        Returns
        -------
        errorCode : integer 
            0 if 'OK', -1 if error  

        Notes
        -----
        The ASCII file should have .DAT extension for sequential objects. This 
        is generally used to specifiy uniform array sag/phase data for 
        Grid Sag / Grid Phase surfaces. 
        """
        cmd = "ImportExtraData,{:d},{}".format(surfNum, fileName)
        reply = self._sendDDEcommand(cmd)
        if 'OK' in reply.rstrip():
            return 0
        else:
            return -1

    def zInsertConfig(self, configNum):
        """Insert a new configuration (column) in the multi-configuration
        editor.

        The new configuration will be placed at the location (column)
        indicated by the parameter ``configNum``

        Parameters
        ----------
        configNum : integer
            the configuration (column) number to insert.

        Returns
        -------
        configCol : integer
            the column number of the configuration that inserted at
            ``configNum``

        Notes
        -----
        1. The ``configNum`` returned (configCol) is generally different
           from the number in the input ``configNum``.
        2. Use ``zInsertMCO()`` to insert a new multi-configuration operand
           in the multi-configuration editor.
        3. Use ``zSetConfig()`` to switch the current configuration number

        See Also
        --------
        zDeleteConfig()
        """
        return int(self._sendDDEcommand("InsertConfig,{:d}".format(configNum)))

    def zInsertMCO(self, operNum):
        """Insert a new multi-configuration operand (row) in the multi-
        configuration editor.

        Parameters
        ----------
        operNum : integer
            number between 1 and the current number of operands plus 1,
            inclusive.

        Returns
        -------
        numOper : integer
            new number of operands (rows)

        See Also
        --------
        zInsertConfig() :
            to insert a new configuration (row)
        zDeleteMCO()
        """
        return int(self._sendDDEcommand("InsertMCO,{:d}".format(operNum)))

    def zInsertMFO(self, operNum):
        """Insert a new optimization operand (row) in the merit function
        editor.

        Parameters
        ----------
        operNum : integer
            number between 1 and the current number of operands plus 1,
            inclusive.

        Returns
        -------
        numOper : integer
            new number of operands (rows).

        See Also
        --------
        zSetOperand() :
            Generally, you may want to use ``zSetOperand()`` afterwards.
        zDeleteMFO()
        """
        return int(self._sendDDEcommand("InsertMFO,{:d}".format(operNum)))

    def zInsertObject(self, surfNum, objNum):
        """Insert a new NSC object at the location indicated by the parameters 
        ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer
            object number

        Returns
        -------
        status : integer
            0 if successful, -1 if failed.

        See Also
        --------
        zSetNSCObjectData() :
            to define data for the new surface
        zDeleteObject()
        """
        cmd = "InsertObject,{:d},{:d}".format(surfNum,objNum)
        reply = self._sendDDEcommand(cmd)
        if reply.rstrip() == 'BAD COMMAND':
            return -1
        else:
            return int(reply.rstrip())

    def zInsertSurface(self, surfNum):
        """Insert a surface at the location indicated by ``surfNum``

        Parameters
        ----------
        surfNum  : integer
            location where to insert the surface

        Returns
        -------
        status : integer
            0 if success

        See Also
        --------
        zSetSurfaceData() :
            to define data for the new surface
        zDeleteSurface()
        """
        return int(self._sendDDEcommand("InsertSurface,"+str(surfNum)))

    def zLoadDetector(self, surfNum, objNum, fileName):
        """Loads the data saved in a file to an NSC Detector Rectangle,
        Detector Color, Detector Polar, or Detector Volume object.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer
            object number
        fileName : string
            the filename may include the full path; if no path is provided
            the path of the current lens file is used. The extension should
            be DDR, DDC, DDP, or DDV for Detector Rectangle, Color, Polar,
            and Volume objects, respectively

        Returns
        -------
        status : integer
            0 if load was successful; Error code (such as -1,-2) if failed.
        """
        isRightExt = _os.path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not _os.path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        isFile = _os.path.isfile(fileName)  # check if file exist
        if isRightExt and isFile:
            cmd = ("LoadDetector,{:d},{:d},{}"
                   .format(surfNum,objNum,fileName))
            reply = self._sendDDEcommand(cmd)
            return _regressLiteralType(reply.rstrip())
        else:
            return -1

    def zLoadFile(self, fileName, append=None):
        """Loads a ZEMAX file into the server

        Parameters
        ----------
        filename : string
            full path of the ZEMAX file to be loaded.
        append : integer, optional
            if a non-zero value of append is passed, then the new file is
            appended to the current file starting at the surface number
            defined by the value appended.

        Returns
        -------
        status : integer
            0 = file successfully loaded;
            -999 = file could not be loaded (check if the filename pattern is 
                problematic or check the path);
            -998 = the command timed out;
            other = the upload failed.

        Notes
        -----
        Filename patterns that are fine:

            a. "C:\\ZEMAX\\Samples\\cooke.zmx"
            b. "C:\ZEMAX\Samples\cooke.zmx"
            c. "C:\\ZEMAX\\My Documents\\Sample\\cooke.zmx"  # spaces in file 
               path is OK.

        Problematic filename patterns:

            a. "C:\ZEMAX\Samples\Example, cooke.zmx"   # a comma (,) in the 
               filename is problematic.

        Examples
        -------- 
        >>> lens = "C:\ZEMAX\Samples\cooke.zmx"
        >>> ln.zLoadFile(lens)
        0
        >>> lens = os.path.join(ln.zGetPath()[1], 'Sequential', 'Objectives', 
                               'Cooke 40 degree field.zmx')
        >>> ln.zLoadFile(lens)
        0
        >>> usr = os.path.expandvars("%userprofile%")
        >>> zmf = 'Double Gauss 5 degree field.zmx'
        >>> lens = os.path.join(usr, 'Documents\Zemax\Samples\Sequential\Objectives', zmf)
        0

        See Also
        --------
        zSaveFile(), zGetPath(), zPushLens()
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
        """Loads a Zemax .MF or .ZMX file and extracts the merit function
        and places it in the lens loaded in the server.

        Parameters
        ----------
        fileName : string
            name of the merit function file with full path and extension.

        Returns
        -------
        number : integer
            number of operands in the merit function
        merit : float
            merit value of the merit function.

        Returns -999 if function failed.

        Notes
        -----
        1. If the merit function value is 9.00e+009, the merit function
           cannot be evaluated.
        2. Loading a merit function file does not change the data displayed
           in the LDE; the server process has a separate copy of the lens
           data.

        See Also
        --------
        zOptimize(), zSaveMerit()
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
        """Loads a tolerance file previously saved with ``zSaveTolerance``
        and places the tolerances in the lens loaded in the DDE server.

        Parameters
        ----------
        fileName : string
            file name of the tolerance file. If no path is provided in the
            filename, the <data>\Tolerance folder is assumed.

        Returns
        -------
        numTolOperands : integer
            number of tolerance operands loaded;
            -999 if file does not exist
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

    def zMakeGraphicWindow(self, fileName, moduleName, winTitle, textFlag,
                           settingsData=None):
        """Notifies Zemax that graphic data has been written to a file and
        may now be displayed as a Zemax child window.

        The primary purpose of this function is to implement user defined
        features in a client application that look and act like native
        Zemax features.

        Parameters
        ----------
        fileName : string
            the full path and file name to the temporary file that holds
            the graphic data. This must be the same name as passed to the
            client executable in the command line arguments, if any.
        moduleName : string
            the full path and executable name of the client program that
            created the graphic data.
        winTitle : string
            the string which defines the title Zemax should place in the
            top bar of the window.
        textFlag : integer (0 or 1)
            1 = the client can also generate a text version of the data.
                Since the current data is a graphic display (it must be if
                the function is ``zMakeGraphicWindow``) Zemax wants to know
                if the "Text" menu option should be available to the user,
                or if it should be grayed out;
            0 = Zemax will gray out the "Text" menu option and will not
                attempt to ask the client to generate a text version of the
                data;
        settingsData : string
            the settings data is a string of values delimited by spaces
            which are used by the client to define how the data was
            generated. These values are only used by the client, not by
            Zemax. The settings data string holds the options and data that
            would normally appear in a Zemax "settings" style dialog box.
            The settings data should be used to recreate the data if
            required. Because the total length of a data item cannot exceed
            255 characters, the function ``zSetSettingsData()`` may be used
            prior to the call to ``zMakeGraphicWindow()`` to specify the
            settings data string rather than including the data as part of
            ``zMakeGraphicWindow()``. See "How Zemax calls the client" in
            the manual for more details on the settings data.

        Returns
        -------
        None

        Notes
        -----
        There are two ways of using this command:

        >>> ln.zMakeGraphicWindow(fileName, moduleName, winTitle, textFlag, settingsData)

          OR

        >>> ln.zSetSettingsData(0, settingsData)
        >>> ln.zMakeGraphicWindow(fileName, moduleName, winTitle, textFlag)

        Examples
        --------
        A sample item string might look like the following:

        >>> ln.zMakeGraphicWindow('C:\TEMP\ZGF001.TMP',
                                  'C:\ZEMAX\FEATURES\CLIENT.EXE',
                                  'ClientWindow', 1, "0 1 2 12.55")

        This call indicates that Zemax should open a graphic window,
        display the data stored in the file 'C:\TEMP\ZGF001.TMP', and that
        any updates or setting changes can be made by calling the client
        module 'C:\ZEMAX\FEATURES\CLIENT.EXE'. This client can generate a
        text version of the graphic, and the settings data string (used
        only by the client) is "0 1 2 12.55".
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
        """Notifies Zemax that text data has been written to a file and may
        now be displayed as a Zemax child window.

        The primary purpose of this item is to implement user defined
        features in a client application, that look and act like native
        Zemax features.

        Parameters
        ----------
        fileName : string
            the full path and file name to the temporary file that holds
            the text data. This must be the same name as passed to the
            client executable in the command line arguments, if any.
        moduleName : string
            the full path and executable name of the client program that
            created the text data.
        winTitle : string
            the string which defines the title Zemax should place in the
            top bar of the window.
        settingsData : string
            the settings data is a string of values delimited by spaces
            which are used by the client to define how the data was
            generated. These values are only used by the client, not by
            Zemax. The settings data string holds the options and data that
            would normally appear in a Zemax "settings" style dialog box.
            The settings data should be used to recreate the data if
            required. Because the total length of a data item cannot exceed
            255 characters, the function ``zSetSettingsData()`` may be used
            prior to the call to ``zMakeTextWindow()`` to specify the
            settings data string rather than including the data as part of
            ``zMakeTextWindow()``. See "How Zemax calls the client" in the
            manual for more details on the settings data.


        Notes
        -----
        There are two ways of using this command:

        >>> ln.zMakeTextWindow(fileName, moduleName, winTitle, settingsData)

          OR

        >>> ln.zSetSettingsData(0, settingsData)
        >>> ln.zMakeTextWindow(fileName, moduleName, winTitle)


        Examples
        --------
        >>> ln.zMakeTextWindow('C:\TEMP\ZGF002.TMP',
                               'C:\ZEMAX\FEATURES\CLIENT.EXE',
                                'ClientWindow',"6 5 4 12.55")

        This call indicates that Zemax should open a text window, display
        the data stored in the file 'C:\TEMP\ZGF002.TMP', and that any
        updates or setting changes can be made by calling the client module
        'C:\ZEMAX\FEATURES\CLIENT.EXE'. The settingsdata string (used only
        by the client) is "6 5 4 12.55".
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
        """Change specific options in Zemax configuration files (.CFG)

        Settings files are used by Zemax analysis windows. They are also
        used by ``zMakeTextWindow()`` and ``zMakeGraphicWindow()``. The
        modified settings file is written back to the original settings
        file-name.

        Parameters
        ----------
        fileName : string
            full name of the settings file, including the path & extension
        mType : string
            a mnemonic that indicates which setting within the file is to
            be modified. See the ZPL macro command "MODIFYSETTINGS" in the
            Zemax manual for a complete list of the ``mType`` codes
        value : string or integer
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
        """Erases the current lens

        The "minimum" lens that remains is identical to the LDE when
        "File >> New" is selected. No prompt to save the existing lens is
        given.

        Parameters
        ----------
        None

        Returns
        -------
        status : integer
            0 = successful
        """
        return int(self._sendDDEcommand('NewLens'))

    def zNSCCoherentData(self, surfNum, detectNum, pixel, dtype):
        """Return data from an NSC detector (Non-sequential coherent data)

        Similar to NSDC optimization operand

        Parameters
        ----------
        surfNum : integer
            the surface number of the NSC group (1 for pure NSC systems).
        detectNum : integer
            the object number of the desired detector.
        pixel : integer
            0 = the sum of the data for all pixels for that detector
                object is returned;
            +ve int = the data from the specified pixel is returned.
        dtype : integer
            0 = real; 1 = imaginary; 2 = amplitude; 3 = power

        Returns
        -------
        nsccoherentdata : float 
            nsc coherent data
        """
        cmd = ("NSCCoherentData,{:d},{:d},{:d},{:d}"
               .format(surfNum, detectNum, pixel, dtype))
        reply = self._sendDDEcommand(cmd)
        return float(reply.rstrip())

    def zNSCDetectorData(self, surfNum, detectNum, pixel, dtype):
        """Return data from an NSC detector (Non-sequential incoherent
        intensity data, similar to NSDD operand)

        Parameters
        ----------
        surfNum : integer
            the surface number of the NSC group (1 for pure NSC systems).
        detectNum : integer
            the object number of the desired detector.
            0 = all detectors are cleared;
            -ve int = only the detector defined by the absolute value of
            ``detectNum`` is cleared
        pixel : integer
            the ``pixel`` argument is interpreted differently depending
            upon the type of detector as follows:

            1. For Detector Rectangles, Detector Surfaces, & all faceted
               detectors (type 1):

                * +ve int = the data from the specified pixel is returned.
                * 0 = the sum of the total flux in position space, average
                  flux/area in position space, or total flux in angle
                  space for all pixels for that detector object, for
                  Data = 0, 1, or 2, respectively.
                * -1 = Maximum flux or flux/area.
                * -2 = Minimum flux or flux/area.
                * -3 = Number of rays striking the detector.
                * -4 = Standard deviation (RMS from the mean) of all the
                  non-zero pixel data.
                * -5 = The mean value of all the non-zero pixel data.
                * -6,-7,-8 = The x, y, or z coordinate of the position or
                  angle Irradiance or Intensity  centroid, resp.
                * -9,-10,-11,-12,-13 = The RMS radius, x, y, z, or xy
                  cross term distance or angle of all the pixel data
                  with respect to the centroid. These are the second
                  moments r^2, x^2, y^2, z^2, & xy, respectively.

            2. For Detector volumes (type 2) ``pixel`` is interpreted as
               the voxel number. if ``pixel==0``,  the value returned
               is the sum for all pixels.

        dtype : integer
            0 = flux or incident flux for type 1 and type 2 detectors
                respectively.
            1 = flux/area for type 1 detectors. Equals absorbed flux for
                type 2 detectors.
            2 = flux/solid angle pixel for type 1 detectors. Equals
                absorbed flux per unit volume for type 2 detectors.

        Notes
        -----
        Only ``dtype`` values 0 & 1 (for flux & flux/area) are supported
        for faceted detectors.
        """
        cmd = ("NSCDetectorData,{:d},{:d},{:d},{:d}"
               .format(surfNum, detectNum, pixel, dtype))
        reply = self._sendDDEcommand(cmd)
        return float(reply.rstrip())

    def zNSCLightningTrace(self, surfNum, source, raySampling, edgeSampling,
                           timeout=60):
        """Traces rays from one or all NSC sources using Lighting Trace

        Parameters
        ----------
        surfNum : integer
            surface number. use 1 for pure NSC mode
        source : integer
            object number of the desired source. If ``0``, all sources
            will be traced.
        raySampling : integer
            resolution of the LightningTrace mesh with valid values
            between 0 (= "Low (1X)") and 5 (= "1024X").
        edgeSampling : integer
            resolution used in refining the LightningTrace mesh near
            the edges of objects, with valid values between 0 ("Low (1X)")
            and 4 ("256X").
        timeout : integer
            timeout value in seconds. Default=60sec

        Returns
        -------
        status : integer

        Notes
        -----
        ``zNSCLightningTrace()`` always updates the lens before executing
        a LightningTrace to make certain all objects are correctly loaded
        and updated.
        """
        cmd = ("NSCLightningTrace,{:d},{:d},{:d},{:d}"
               .format(surfNum, source, raySampling, edgeSampling))
        reply = self._sendDDEcommand(cmd, timeout)
        if 'OK' in reply.split():
            return 0
        elif 'BAD COMMAND' in reply.rstrip():
            return -1
        else:
            return int(float(reply.rstrip()))  # return the error code sent by zemax.

    def zNSCTrace(self, surfNum, srcNum, split=0, scatter=0, usePolar=0,
                  ignoreErrors=0, randomSeed=0, save=0, saveFilename=None,
                  oFilter=None, timeout=180):
        """Trace rays from one or all NSC sources, after updating the lens.

        Parameters
        ----------
        surfNum : integer
            the surface number of the NSC group (1 for pure NSC systems).
        srcNum : integer
            the object number of the source. Use 0 to trace all sources.
        split : integer, optional
            0 = splitting is OFF (default); otherwise splitting is ON
        scatter : integer, optional
            0 = scattering is OFF (default); otherwise scattering is ON
        usePolar : integer
            0 = polarization is OFF (default); otherwise polarization is
            ON. If splitting is ON polarization is automatically selected.
        ignoreErrors : integer, optional
            0 = ray errors will terminate the NSC trace & macro execution
            and an error will be reported (default). Otherwise errors will
            be ignored
        randomSeed : integer, optional
            0 or omitted = the random number generator will be seeded with
            a random value, & every call to ``zNSCTrace()`` will produce
            different random rays (default). Any integer other than zero
            will ensure that the random number generator be seeded with
            the specified value, and every call to ``zNSCTrace()`` using
            the same seed will produce identical rays.
        save : integer, optional
            0 or omitted = the parameters ``saveFilename`` and ``oFilter``
            need not be supplied (default). Otherwise the rays will be
            saved in a ``ZRD`` file. The ``ZRD`` file will have the name
            specified by the ``saveFilename``, and will be placed in the
            same directory as the lens file. The extension of
            ``saveFilename`` should be ``ZRD``, and no path should be
            specified.
        saveFilename : string, optional
            (see above)
        oFilter : string, optional
            if ``save`` is not zero, then the optional filter name is
            either a string variable with the filter, or the literal
            filter in double quotes. For information on filter strings
            see "The filter string" in the Zemax manual.
        timeout : integer
            timeout in seconds (default = 60 seconds)

        Returns
        -------
        traceResult : error code
            0 if successful, -1 if problem with saveFileName, other
            error codes sent by Zemax.

        Examples
        --------
        >>> zNSCTrace(1, 2, 1, 0, 1, 1)

        The above command traces rays in NSC group 1, from source 2, with
        ray splitting, no ray scattering, using polarization and ignoring
        errors.

        >>> zNSCTrace(1, 2, 1, 0, 1, 1, 33, 1, "myrays.ZRD", "h2")

        Same as above, only a random seed of 33 is given and the data is
        saved to the file "myrays.ZRD" after filtering as per h2.

        >>> zNSCTrace(1, 2)

        The above command traces rays in NSC group 1, from source 2,
        without ray splitting, no ray scattering, without using
        polarization and will not ignore errors.

        See Also
        -------- 
        zNSCDetectorClear()
        """
        requiredArgs = ("{:d},{:d},{:d},{:d},{:d},{:d},{:d},{:d}"
        .format(surfNum,srcNum, split, scatter, usePolar, ignoreErrors,
                randomSeed, save))
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
            return int(float(reply.rstrip()))  # return error code sent by zemax.

    def zOpenWindow(self, analysisType, zplMacro=False, timeout=None):
        """Open a new analysis window in the main Zemax application screen.

        Parameters
        ----------
        analysisType : string
            the 3-letter button code corresponding to the analysis. A list
            of these codes can be seen by calling ``pyz.showZButtons()``
            function in an interactive shell.
        zplMacro : bool, optional
            ``True`` if the ``analysisType`` code is the first 3-letters
            of a ZPL macro name, else ``False`` (default).
        timeout : integer, optional
            timeout value in seconds.

        Returns
        -------
        status : integer
            0 = successful; -1 = incorrect analysis code; -999 = fail

        Notes
        -----
        1. This function checks if the ``analysisType`` code is a valid
           code or not in order to prevent the calling program from
           getting stalled. However, it doesn't check the ``analysisType``
           code validity if it is a ZPL macro. If the ``analysisType`` is
           ZPL macro, please make sure that the macro exist in the
           ``<data>/Macros`` folder.
        2. You may also use ``zExecuteZPLMacro()`` to execute a ZPL macro.

        See Also
        --------
        zGetMetaFile(), zExecuteZPLMacro()
        """
        if zb.isZButtonCode(analysisType) ^ zplMacro:
            reply = self._sendDDEcommand("OpenWindow,{}".format(analysisType),
                                         timeout)
            if 'OK' in reply.split():
                return 0
            elif 'FAIL' in reply.split():
                return -999
            else:
                return int(float(reply.rstrip()))  # error code from Zemax
        else:
            return -1

    def zOperandValue(self, operandType, *values):
        """Returns the value of any optimization operand, even if the
        operand is not currently in the merit function.

        Parameters
        ----------
        operandType : string
            a valid optimization operand
        *values : flattened sequence
            a sequence of arguments. Possible arguments include:
            ``int1`` (column 2, integer), ``int2`` (column 3, integer),
            ``data1`` (column 4, float), ``data2`` (column 5, float),
            ``data3`` (column 6, float), ``data4`` (column 7, float),
            ``data5`` (column 12, float), ``data6`` (column 13, float)

        Returns
        -------
        operandValue : float
            the value of the operand

        Examples
        --------
        The following example retrieves the total optical path length
        of the marginal ray between surfaces 1 and 3

        >>> ln.zOperandValue('PLEN', 1, 3, 0, 0, 0, 1)

        See Also
        --------
        zOptimize():
            to update MFE prior to calling ``zOperandValue()``, call
            ``zOptimize(-1)``
        zGetOperand(), zSetOperand()
        """
        if zo.isZOperand(operandType, 1) and (0 < len(values) < 9):
            valList = [str(int(elem)) if i in (0,1) else str(float(elem))
                       for i,elem in enumerate(values)]
            arguments = ",".join(valList)
            cmd = "OperandValue," + operandType + "," + arguments
            reply = self._sendDDEcommand(cmd)
            return float(reply.rstrip())
        else:
            return -1

    def zOptimize(self, numOfCycles=0, algorithm=0, timeout=None):
        """Calls Damped Least Squares/ Orthogonal Descent optimizer.

        Parameters
        ----------
        numOfCycles : integer, optional
            the number of cycles to run. If ``numOfCycles == 0``
            (default), optimization runs in automatic mode. Else if
            ``numOfCycles < 0``, ``zOptimize()`` updates all operands
            in the merit function and returns the current merit function
            without performing optimization.
        algorithm : integer, optional
            0 = Damped Least Squares; 1 = Orthogonal descent
        timeout : integer, optional
            timeout value in seconds

        Returns
        -------
        finalMeritFn : float
            the final merit function.

        Notes
        -----
        1. If the merit function value returned is 9.0E+009, the
           optimization failed, usually because the lens or merit function
           could not be evaluated.
        2. The number of cycles should be kept small enough to allow the
           algorithm to complete and return before the DDE communication
           times out, or an error will occur. One possible way to achieve
           high number of cycles could be to call ``zOptimize()`` multiple
           times in a loop, each time comparing the returned merit
           function with few of the previously returned (and stored) merit
           function values to determine if an optimum has been attained.
           For an example implementation, see ``zOptimize2()``

        See Also
        --------
        zHammer(), zLoadMerit(), zSaveMerit(), zOptimize2()
        """
        cmd = "Optimize,{:1.2g},{:d}".format(numOfCycles,algorithm)
        reply = self._sendDDEcommand(cmd, timeout)
        return float(reply.rstrip())

    def zPushLens(self, update=None, timeout=None):
        """Copy lens in the Zemax DDE server into Lens Data Editor (LDE).

        Parameters
        ----------
        update : integer, optional
            if 0 or omitted, the open windows in Zemax main application
            are not updated;
            if 1, then all open analysis windows are updated.
        timeout : integer, optional
            if a timeout, in seconds, in passed, the client will wait till
            the timeout before returning a timeout error. If no timeout is
            passed, the default timeout is used.

        Returns
        -------
        status : integer
            0 = lens successfully pushed into the LDE;
            -999 = the lens could not be pushed into the LDE. (check
                   ``zPushLensPermission()``);
            -998 = the command timed out;
             other = the update failed.

        Notes
        -----
        This operation requires the permission of the user running the
        Zemax program. The proper use of ``zPushLens`` is to first call
        ``zPushLensPermission()``.

        See Also
        --------
        zPushLensPermission(), zLoadFile(), zGetUpdate(), zGetPath(),
        zGetRefresh(), zSaveFile().
        """
        reply = None
        if update == 1:
            reply = self._sendDDEcommand('PushLens,1', timeout)
        elif update == 0 or update is None:
            reply = self._sendDDEcommand('PushLens,0', timeout)
        else:
            raise ValueError('Invalid value for flag')
        if reply:
            return int(reply)   # Note: Zemax returns -999 if push lens fails
        else:
            return -998         # if timeout reached (assumption!!)

    def zPushLensPermission(self):
        """Establish if Zemax extensions are allowed to push lenses in
        the LDE.

        Parameters
        ----------
        None

        Returns
        -------
        status : integer
            1 = Zemax is set to accept PushLens commands;
            0 = Extensions are not allowed to use ``zPushLens()``

        See Also
        --------
        zPushLens(), zGetRefresh()
        """
        status = None
        status = self._sendDDEcommand('PushLensPermission')
        return int(status)

    def zQuickFocus(self, mode=0, centroid=0):
        """Quick focus adjustment of back focal distance for best focus

        The "best" focus is chosen as a wavelength weighted average over
        all fields.

        Parameters
        ----------
        mode : integer, optional
            0 = RMS spot size (default)
            1 = spot x
            2 = spot y
            3 = wavefront OPD
        centroid : integer, optional
            specify RMS reference
             0 = RMS referenced to the chief ray (default);
             1 = RMS referenced to image centroid

        Returns
        -------
        retVal : integer
            0 for success.
        """
        retVal = -1
        cmd = "QuickFocus,{mode:d},{cent:d}".format(mode=mode,cent=centroid)
        reply = self._sendDDEcommand(cmd)
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zReleaseWindow(self, tempFileName):
        """Release locked window/menu mar.

        Parameters
        ----------
        tempFileName : string
            the temporary file name

        Returns
        -------
        status : integer
            0 = no window is using the filename;
            1 = the file is being used.

        Notes
        -----
        When Zemax calls the client to update or change the settings used
        by the client function, the menu bar is grayed out on the window
        to prevent multiple updates or setting changes from being
        requested simultaneously. Normally, when the client code calls
        the functions ``zMakeTextWindow()`` or ``zMakeGraphicWindow()``,
        the menu bar is once again activated. However, if during an update
        or setting change, the new data cannot be computed, then the
        window must be released. The ``zReleaseWindow()`` function serves
        just this one purpose. If the user selects "Cancel" when changing
        the settings, the client code should send a ``zReleaseWindow()``
        call to release the lock out of the menu bar. If this command is
        not sent, the window cannot be closed, which will prevent
        Zemax from terminating normally.
        """
        reply = self._sendDDEcommand("ReleaseWindow,{}".format(tempFileName))
        return int(float(reply.rstrip()))

    def zRemoveVariables(self):
        """Sets all currently defined solve variables to fixed status

        Parameters
        ----------
        None

        Returns
        -------
        status : integer
            0 = successful; -1 = fail
        """
        reply = self._sendDDEcommand('RemoveVariables')
        if 'OK' in reply.split():
            return 0
        else:
            return -1

    def zSaveDetector(self, surfNum, objNum, fileName):
        """Saves the data currently on an NSC Detector Rectangle, Detector
        Color, Detector Polar, or Detector Volume object to a file.

        Parameters
        ----------
        surfNum : integer
            surface number of the non-sequential group. Use 1 if the
            program mode is Non-Sequential.
        objNum : integer
            object number of the detector object
        fileName : string
            the filename may include the full path; if no path is provided
            the path of the current lens file is used. The extension should
            be DDR, DDC, DDP, or DDV for Detector Rectangle, Color, Polar,
            and Volume objects, respectively.

        Returns
        -------
        status : integer
            0 if save was successful;
            Error code (such as -1,-2) if failed.
        """
        isRightExt = _os.path.splitext(fileName)[1] in ('.ddr','.DDR','.ddc','.DDC',
                                                    '.ddp','.DDP','.ddv','.DDV')
        if not _os.path.isabs(fileName): # full path is not provided
            fileName = self.zGetPath()[0] + fileName
        if isRightExt:
            cmd = ("SaveDetector,{:d},{:d},{}"
                   .format(surfNum, objNum, fileName))
            reply = self._sendDDEcommand(cmd)
            return _regressLiteralType(reply.rstrip())
        else:
            return -1

    def zSaveFile(self, fileName):
        """Saves the lens currently loaded in the server to a Zemax file.

        Parameters
        ----------
        fileName : string
            file name, including full path with extension.

        Returns
        -------
        status : integer
            0 = Zemax successfully saved the lens file & updated the
            newly saved lens;
            -999 = Zemax couldn't save the file;
            -1 = Incorrect file name;
            Any other value = update failed.

        See Also
        --------
        zGetPath(), zGetRefresh(), zLoadFile(), zPushLens().
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
        """Saves the current merit function to a Zemax .MF file

        Parameters
        ----------
        fileName : string
            name of the merit function file with full path and extension.

        Returns
        -------
        meritData : integer
            If successful, it is the number of operands in the merit
            function; If ``meritData = -1``, saving failed.

        See Also
        --------
        zOptimize(), zLoadMerit()
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

        Parameters
        ----------
        fileName : string
            filename of the file to save the tolerance data. If no path
            is provided, the ``<data>\Tolerance`` folder is assumed.
            Although it is not enforced, it is useful to use ".tol" as
            extension.

        Returns
        -------
        numTolOperands : integer
            number of tolerance operands saved.

        See Also
        --------
        zLoadTolerance()
        """
        cmd = "SaveTolerance,{}".format(fileName)
        reply = self._sendDDEcommand(cmd)
        return int(float(reply.rstrip()))

    def zSetAperture(self, surf, aType, aMin, aMax, xDecenter=0, yDecenter=0,
                     apertureFile=''):
        """Set aperture characteristics at a lens surface (surface data
        dialog box)

        Parameters
        ----------
        surf : integer
            surface number
        aType : integer
            code to specify aperture type
                * 0 = no aperture (na)
                * 1 = circular aperture (ca)
                * 2 = circular obscuration (co)
                * 3 = spider (s)
                * 4 = rectangular aperture (ra)
                * 5 = rectangular obscuration (ro)
                * 6 = elliptical aperture (ea)
                * 7 = elliptical obscuration (eo)
                * 8 = user defined aperture (uda)
                * 9 = user defined obscuration (udo)
                * 10 = floating aperture (fa)
        aMin : float
            min radius (ca), min radius (co), width of arm (s), x-half
            width (ra), x-half width (ro), x-half width (ea), x-half
            width (eo)
        aMax : float
            max radius (ca), max radius (co), number of arm (s),
            y-half width (ra), y-half width (ro), y-half width (ea),
            y-half width(eo). See [AT]_ for details.
        xDecenter : float, optional
            amount of decenter from current optical axis (lens units)
        yDecenter : float, optional
            amount of decenter from current optical axis (lens units)
        apertureFile : string, optional
            a text file with .UDA extention. See [UDA]_ for detils.

        Returns
        -------
        apertureInfo : tuple
            apertureInfo is a tuple containing the following:
                * aType : (see above)
                * aMin  : (see above)
                * aMax  : (see above)
                * xDecenter : (see above)
                * yDecenter : (see above)

        Examples
        --------
        >>> apertureInfo = ln.zSetAperture(2, 1, 5, 10, 0.5, 0, 'apertureFile.uda')
        or
        >>> apertureInfo = ln.zSetAperture(2, 1, 5, 10)

        Notes
        -----
        1. The ``aMin`` and ``aMax`` values have different meanings for
           the elliptical, rectangular, and spider apertures than for
           circular apertures
        2. If ``zSetAperture()`` is used to set user defined apertures
           or obscurations, the ``aperturefile`` must be the name of a
           file which lists the x, y, coordinates of the user defined
           aperture file in a two column format. For more information
           on user defined apertures, see [UDA]_

        References
        ----------
        .. [AT] "Aperture type and other aperture controls," Zemax manual.
        .. [UDA]  "User defined apertures and obscurations," Zemax manual.

        See Also
        --------
        zGetAperture()
        """
        cmd  = ("SetAperture,{sN:d},{aT:d},{aMn:1.20g},{aMx:1.20g},{xD:1.20g},"
                "{yD:1.20g},{aF}".format(sN=surf, aT=aType, aMn=aMin, aMx=aMax,
                 xD=xDecenter, yD=yDecenter, aF=apertureFile))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        ainfo = _co.namedtuple('ApertureInfo', ['aType', 'aMin', 'aMax',
                                                'xDecenter', 'yDecenter'])
        apertureInfo = ainfo._make([float(elem) for elem in rs])
        return apertureInfo

    def zSetBuffer(self, bufferNum, textData):
        """Used to store client specific data with the window being
        created or updated.

        The buffer data can be used to store user selected options
        instead of using the settings data on the command line of the
        ``zMakeTextWindow()`` or ``zMakeGraphicWindow()`` functions.
        The data must be in a string format.

        Parameters
        ----------
        bufferNum : integer
            number between 0 and 15 inclusive (for 16 buffers provided)
        textData : string
            is the only text that is stored, maximum of 240 characters

        Returns
        -------
        status : integer
            0 if successful, else -1

        Notes
        -----
        The buffer data is not associated with any particular window until
        either the ``zMakeTextWindow()`` or ``zMakeGraphicWindow()``
        functions are issued. Once Zemax receives the ``MakeTextWindow``
        or ``MakeGraphicWindow`` items, the buffer data is then copied to
        the appropriate window memory, and then may later be retrieved
        from that window's buffer using ``zGetBuffer()`` function.

        See Also
        --------
        zGetBuffer()
        """
        if (0 < len(textData) < 240) and (0 <= bufferNum < 16):
            cmd = "SetBuffer,{:d},{}".format(bufferNum, str(textData))
            reply = self._sendDDEcommand(cmd)
            return 0 if 'OK' in reply.rsplit() else -1
        else:
            return -1

    def zSetConfig(self, config):
        """Switches the current configuration number (selected column in
        the MCE), and updates the system.

        Parameters
        ----------
        config : integer
            The configuration (column) number to set current

        Returns
        -------
        currentConfig : integer
            the current configuration (column) number in MCE
            ``1 <= currentConfig <= numberOfConfigs``
         numberOfConfigs : integer
            number of configurations (columns)
         error : integer
            0 = successful (i.e. new current config is traceable);
            -1 = failure

        Notes
        -----
        Use ``zInsertConfig()`` to insert new configuration in the
        multi-configuration editor.

        See Also
        --------
        zGetConfig(), zSetMulticon()
        """
        reply = self._sendDDEcommand("SetConfig,{:d}".format(config))
        rs = reply.split(',')
        return tuple([int(elem) for elem in rs])

    def zSetExtra(self, surfNum, colNum, value):
        """Sets extra surface data (value) in the Extra Data Editor for
        the surface indicatd by ``surf``

        Parameters
        ----------
        surfNum : integer
            the surface number
        colNum : integer
            the column number
        value : float
            the value

        Returns
        -------
        retValue : float
            the numeric data value

        See Also
        --------
        zGetExtra()
        """
        cmd = ("SetExtra,{:d},{:d},{:1.20g}".format(surfNum, colNum, value))
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zSetField(self, n, arg1, arg2, arg3=None, vdx=0.0, vdy=0.0,
                  vcx=0.0, vcy=0.0, van=0.0):
        """Sets the field data for a particular field point

        There are 2 ways of using this function (the parameters ``arg1``,
        ``arg2`` and ``arg3`` have different meanings depending on ``n``):

            * ``zSetField(0, fieldType, totalNumFields, normMethod)``

             OR

            * ``zSetField(n, xf, yf [,wgt, vdx, vdy, vcx, vcy, van])``

        Parameters
        ----------
        [Case: ``n = 0``]

        n : 0
            to set general field parameters
        arg1 : integer
            the field type. 0 = angle, 1 = object height, 2 = paraxial
            image height, and 3 = real image height
        arg2 : integer
            total number of fields
        arg3 : integer (0 or 1), optional
            normalization type. 0 = radial (default), 1 = rectangular

        [Case: ``0 < n <= number_of_fields``]

        n : integer (greater than 0)
            the field number
        arg1 (fx) : float
            the field-x value
        arg2 (fy) : float
            the field-y value
        arg3 (wgt) : float, optional
            the field weight (default = 1.0)
        vdx, vdy, vcx, vcy, van : floats, optional
            the vignetting factors (default = 0.0). See below.

        Returns
        -------
        [Case: ``n=0``]

        type : integer
            0 = angles in degrees; 1 = object height; 2 = paraxial image
            height, 3 = real image height
        number : integer
            number of fields currently defined
        maxX : float
            values used to normalize x field coordinate
        maxY : float
            values used to normalize y field coordinate
        normMethod : integer
            normalization method (0 = radial, 1 = rectangular)

        [Case: ``0 < n <= number-of-fields``]

        xf : float
            the field-x value
        yf : float
            the field-y value
        wgt : float
            field weight value
        vdx : float
            decenter-x vignetting factor
        vdy : float
            decenter-y vignetting factor
        vcx : float
            compression-x vignetting factor
        vcy : float
            compression-y vignetting factor
        van : float
            angle vignetting factor

        Notes
        -----
        1. In Zemax main application, the default field normalization type
           is radial. However, the default field normalization
        2. The returned tuple's content and structure is exactly same as
           that of ``zGetField()``

        See Also
        --------
        zGetField()
        """
        if n:
            fd = _co.namedtuple('fieldData', ['xf', 'yf', 'wgt',
                                              'vdx', 'vdy',
                                              'vcx', 'vcy', 'van'])
            arg3 = 1.0 if arg3 is None else arg3 # default weight
            cmd = ("SetField,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                   ",{:1.20g},{:1.20g},{:1.20g}"
                   .format(n, arg1, arg2, arg3, vdx, vdy, vcx, vcy, van))
        else:
            fd = _co.namedtuple('fieldData', ['type', 'numFields',
                                              'maxX', 'maxY', 'normMethod'])
            arg3 = 0 if arg3 is None else arg3 # default normalization
            cmd = ("SetField,{:d},{:d},{:d},{:d}".format(0, arg1, arg2, arg3))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        if n:
            fieldData = fd._make([float(elem) for elem in rs])
        else:
            fieldData = fd._make([int(elem) if (i==0 or i==1 or i==4)
                                 else float(elem) for i, elem in enumerate(rs)])
        return fieldData

    def zSetFloat(self):
        """Sets all surfaces without surface apertures to have floating
        apertures. Floating apertures will vignette rays which trace
        beyond the semi-diameter.

        Parameters
        ----------
        None

        Returns
        -------
        status : integer
            0 = success; -1 = fail
        """
        retVal = -1
        reply = self._sendDDEcommand('SetFloat')
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zSetLabel(self, surfNum, label):
        """This command associates an integer label with the specified
        surface. The label will be retained by Zemax as surfaces are
        inserted or deleted around the target surface.


        Parameters
        ----------
        surfNum : integer
            the surface number
        label : integer
            the integer label

        Returns
        -------
        assignedLabel : integer
            should be equal to label

        See Also
        --------
        zGetLabel(), zFindLabel()
        """
        reply = self._sendDDEcommand("SetLabel,{:d},{:d}"
                                          .format(surfNum,label))
        return int(float(reply.rstrip()))

    def zSetMulticon(self, config, *multicon_args):
        """Set data or operand type in the multi-configuration editior.

        Note that there are 2 ways of using this function.

        [``USAGE TYPE - I``]

        If ``config > 0``, then the function is used to set data in the
        MCE using the following syntax:

        ``ln.zSetMulticon(config, row, value, status, pickupRow, pickupConfig, scale, offset) -> multiConData``

        Parameters
        ----------
        config : integer (``> 0``)
            configuration number (column)
        row : integer
            the row or operand number
        value : float
            the value to set
        status : integer
            the ``status`` is 0 for fixed, 1 for variable, 2 for pickup,
            and 3 for thermal pickup.
            If ``status`` is 2 or 3, the ``pickupRow`` and ``pickupConfig``
            values indicate the source data for the pickup solve.
        pickupRow : integer
            see ``status``
        pickupConfig : integer
            see ``status``
        scale : float
            scale factor for the pickup value 
        offset : float
            offset to add to the pickup value. 

        Returns
        -------
        multiConData : namedtuple
            the ``multiConData`` is a 8-tuple whose elements are:
            (value, numConfig, numRow, status, pickupRow,
            pickupConfig, scale, offset)


        [``USAGE TYPE - II``]

        If the ``config = 0``, the function may be used to set the operand
        type and number data using the following syntax:

        ``ln.zSetMulticon(0, row, operandType, num1, num2, num3) -> multiConData``

        Parameters
        ----------
        config : 0
            for usage type II
        row : integer
            row or operand number in the MCE
        operandType : string
            the operand type, such as 'THIC', 'WLWT', etc.
        num1 : integer
            number data. `num1` could be "Surface#", "Surface", "Field#", 
            "Wave#', or "Ignored". See [MCO]_
        num2 : integer
            number data. `num2` could be "Object", "Extra Data Number", 
            or "Parameter". See [MCO]_
        num3 : integer
            number data. `num3` could be "Property", or "Face#". See [MCO]_

        Returns
        -------
        multiConData is a 4-tuple (named) whose elements are:
        (operandType, num1, num2, num3)

        Examples
        --------
        The following example shows the USEAGE TYPE - I:

        >>> multiConData = ln.zSetMulticon(1, 5, 5.6, 0, 0, 0, 1.0, 0.0)

        The following two lines show how to set a variable solve on the operand 
        on the 4th row for configuration number 1 (the third line is the output):

        >>> config=1; row=4; value=0.5; status=1; pickupRow=0; pickupConfig=0; scale=1; offset=0
        >>> ln.zSetMulticon(config, row, value, status, pickupRow, pickupConfig, scale, offset)
        MCD(value=0.5, numConfig=2, numRow=4, status=1, pickupRow=0, pickupConfig=0, scale=1.0, offset=0.0)

        The following example shows the USAGE TYPE - II:

        >>> multiConData = ln.zSetMulticon(0, 5, 'THIC', 15, 0, 0)

        Notes
        -----
        1. If there are current operands in the MCE, it is recommended to
           first use ``zInsertMCO()`` to insert a row, and then use
           ``zSetMulticon(0,...)``. For example, use ``zInsertMCO(5)``
           and then use ``zSetMulticon(0, 5, 'THIC', 15, 0, 0)``.
           If a row is not inserted first, then existing rows may be
           overwritten.
        2. The function raises an exception if it determines the arguments
           to be invalid.

        References
        ----------
        .. [MCO] "Summary of Multi-Configuration Operands," Zemax manual.

        See Also
        --------
        zGetMulticon()
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
            mcd = _co.namedtuple('MCD', ['value', 'numConfig', 'numRow', 'status',
                                         'pickupRow', 'pickupConfig', 'scale',
                                         'offset'])
            rs = reply.split(",")
            multiConData = [float(rs[i]) if (i == 0 or i == 6 or i== 7) else int(rs[i])
                                                 for i in range(len(rs))]
        else: # if config == 0
            mcd = _co.namedtuple('MCD', ['operandType', 'num1', 'num2', 'num3'])
            rs = reply.split(",")
            multiConData = [int(elem) for elem in rs[1:]]
            multiConData.insert(0,rs[0])
        return mcd._make(multiConData)

    def zSetNSCObjectData(self, surfNum, objNum, code, data):
        """Sets the various data for NSC objects.

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        code : integer
            integer code
        data : string/integer/float
            data to set NSC object

        Returns
        -------
        nscObjectData : string/integer/float
            the returned data (same as returned by ``zGetNSCObjectData()``)
            depends on the ``code``. If the command fails, it returns ``-1``.
            Refer table nsc-object-data-codes_.

        Notes
        -----
        Refer table nsc-object-data-codes_ in the docstring of
        ``zGetNSCObjectData()`` for ``code`` and ``data`` specific details.

        See Also
        --------
        zGetNSCObjectData(), zSetNSCObjectFaceData()
        """
        str_codes = (0,1,4)
        int_codes = (2,3,5,6,29,101,102,110,111)
        if code in str_codes:
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{}"
              .format(surfNum,objNum,code,data))
        elif code in int_codes:
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{:d}"
              .format(surfNum,objNum,code,data))
        else:  # data is float
            cmd = ("SetNSCObjectData,{:d},{:d},{:d},{:1.20g}"
              .format(surfNum,objNum,code,data))
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

    def zSetNSCObjectFaceData(self, surfNum, objNum, faceNum, code, data):
        """Sets the various data for NSC object faces

        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        faceNum : integer
            face number
        code : integer
            integer code
        data : float/integer/string
            data to set NSC object face

        Returns
        -------
        nscObjFaceData  : string/integer/float
            the returned data (same as returned by ``zGetNSCObjectFaceData()``)
            depends on the ``code``. If the command fails, it returns ``-1``.
            Refer table nsc-object-face-data-codes_.

        Notes
        -----
        Refer table nsc-object-face-data-codes_ in the docstring of
        ``zGetNSCObjectData()`` for ``code`` and ``data`` specific details.

        See Also
        --------
        zGetNSCObjectFaceData()
        """
        str_codes = (10,30,31,40,60)
        int_codes = (20,22,24)
        if code in str_codes:
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{}"
                   .format(surfNum,objNum,faceNum,code,data))
        elif code in int_codes:
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{:d}"
                  .format(surfNum,objNum,faceNum,code,data))
        else: # data is float
            cmd = ("SetNSCObjectFaceData,{:d},{:d},{:d},{:d},{:1.20g}"
                  .format(surfNum,objNum,faceNum,code,data))
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

    def zSetNSCParameter(self, surfNum, objNum, paramNum, data):
        """Sets the parameter data for NSC objects.

        Parameters
        ----------
        surfNum : integer
            the surface number. Use 1 if Non-Sequential program mode
        objNum : integer
            the object number
        paramNum : integer
            the parameter number
        data : float
            the new numeric value for the ``paramNum``

        Returns
        -------
        nscParaVal : float
            the parameter value

        See Also
        --------
        zGetNSCParameter()
        """
        cmd = ("SetNSCParameter,{:d},{:d},{:d},{:1.20g}"
              .format(surfNum, objNum, paramNum, data))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if rs == 'BAD COMMAND':
            nscParaVal = -1
        else:
            nscParaVal = float(rs)
        return nscParaVal

    def zSetNSCPosition(self, surfNum, objNum, code, data):
        """Sets the position data for NSC objects.

        Parameters
        ----------
        surfNum : integer
            the surface number. Use 1 if Non-Sequential program mode
        objNum : integer
            the object number
        code : integer
            1-7 for x, y, z, tilt-x, tilt-y, tilt-z, and material,
            respectively.
        data : float or string
            numeric (float) for codes 1-6, string for material (code-7)

        Returns
        -------
        nscPosData : tuple
            a 7-tuple containing x, y, z, tilt-x, tilt-y, tilt-z, material

        See Also
        --------
        zSetNSCPositionTuple(), zGetNSCPosition()
        """
        if code == 7:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{}"
            .format(surfNum, objNum, code, data))
        else:
            cmd = ("SetNSCPosition,{:d},{:d},{:d},{:1.20g}"
            .format(surfNum, objNum, code, data))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        if rs[0].rstrip() == 'BAD COMMAND':
            nscPosData = -1
        else:
            nscPosData = tuple([str(rs[i].rstrip()) if i==6 else float(rs[i])
                                                    for i in range(len(rs))])
        return nscPosData

    def zSetNSCProperty(self, surfNum, objNum, faceNum, code, value):
        """Sets a numeric or string value to the property pages of objects
        defined in the NSC editor. It mimics the ZPL function NPRO.


        Parameters
        ----------
        surfNum : integer
            surface number of the NSC group. Use 1 if for pure NSC mode
        objNum : integer
            the NSC ojbect number
        faceNum : integer
            face number. Use 0 for "All Faces"
        code : integer
            for the specific code
        value : string/integer/float
            value to set NSC property

        Returns
        -------
        nscPropData : string/float/integer
            the returned data (same as returned by ``zGetNSCProperty()``)
            depends on the ``code``. If the command fails, it returns
            ``-1``. Refer table nsc-property-codes_.

        Notes
        -----
        Refer table nsc-property-codes_ in the docstring of
        ``zGetNSCProperty()`` for ``code`` and ``value`` specific details.

        See Also
        --------
        zGetNSCProperty()
        """
        cmd = ("SetNSCProperty,{:d},{:d},{:d},{:d},".format(surfNum, objNum, code, faceNum))
        if code in (0,1,4,5,6,11,12,14,18,19,27,28,84,86,92,117,123):
            cmd = cmd + value
        elif code in (2,3,7,9,13,15,16,17,20,29,81,91,101,102,110,111,113,121,
                                       141,142,151,152,153161,162,171,172,173):
            cmd = cmd + str(int(value))
        else:
            cmd = cmd + str(float(value))
        reply = self._sendDDEcommand(cmd)
        nscPropData = _process_get_set_NSCProperty(code, reply)
        return nscPropData

    def zSetNSCSettings(self, maxInt, maxSeg, maxNest, minAbsI, minRelI,
                        glueDist, missRayLen, ignoreErr):
        """Sets the maximum number of intersections, segments, nesting
        level, minimum absolute intensity, minimum relative intensity,
        glue distance, miss ray distance, and ignore errors flag used
        for NSC ray tracing.

        Parameters
        ----------
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
            1 if yes, 0 if no

        Returns
        -------
        nscSettingsDataRet : 8-tuple
            the returned tuple is also an 8-tuple with the same elements
            as ``nscSettingsData``.

        Notes
        -----
        Since the ``maxSeg`` value may require large amounts of RAM,
        verify that the new value was accepted by checking the returned
        tuple.

        See Also
        --------
        zGetNSCSettings()
        """
        cmd = ("SetNSCSettings,{:d},{:d},{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g},{:d}"
        .format(maxInt, maxSeg, maxNest, minAbsI, minRelI, glueDist, missRayLen, ignoreErr))
        reply = str(self._sendDDEcommand(cmd))
        rs = reply.rsplit(",")
        nscSettingsData = [float(rs[i]) if i in (3,4,5,6) else int(float(rs[i]))
                                                        for i in range(len(rs))]
        return tuple(nscSettingsData)

    def zSetNSCSolve(self, surfNum, objNum, param, solveType,
                     pickupObject=0, pickupColumn=0, scale=0, offset=0):
        """Sets the solve type on NSC position and parameter data.

        Parameters
        ----------
        surfNum : integer
            the surface number. Use 1 if in Non-Sequential mode.
        objNum : integer
            the object number
        param : integer
            * -1 = data for x position;
            * -2 = data for y position;
            * -3 = data for z position;
            * -4 = data for tilt x ;
            * -5 = data for tilt y ;
            * -6 = data for tilt z ;
            * n > 0  = data for the nth parameter;
        solveType : integer
            0 = fixed; 1 = variable; 2 = pickup;
        pickupObject : integer, optional
            if ``solveType = 2``, pickup object number
        pickupColumn : integer, optional
            if ``solveType = 2``, pickup column number (0 for current column)
        scale : float, optional
            if ``solveType = 2``, scale factor
        offset : float, optional
            if ``solveType = 2``, offset

        Returns
        -------
        nscSolveData : tuple or errorCode
            5-tuple containing
            ``(status, pickupObject, pickupColumn, scaleFactor, offset)``
            The status value is 0 for fixed, 1 for variable, and 2 for
            a pickup solve. Only when the stauts is a pickup solve is the
            other data meaningful.
            -1 if it a BAD COMMAND

        See Also
        --------
        zGetNSCSolve()
        """
        nscSolveData = -1
        args1 = "{:d},{:d},{:d},".format(surfNum, objNum, param)
        args2 = "{:d},{:d},{:d},".format(solveType, pickupObject, pickupColumn)
        args3 = "{:1.20g},{:1.20g}".format(scale, offset)
        cmd = ''.join(["SetNSCSolve,",args1, args2, args3])
        reply = self._sendDDEcommand(cmd)
        rs = reply.rstrip()
        if 'BAD COMMAND' not in rs:
            nscSolveData = tuple([float(e) if i in (3,4) else int(float(e))
                                 for i,e in enumerate(rs.split(","))])
        return nscSolveData

    def zSetOperand(self, row, column, value):
        """Sets the operand data in the Merit Function Editor

        Parameters
        ----------
        row : integer
            operand row number in the MFE
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
        zSetOperandRow():
            sets an entire row of the MFE
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

    def zSetPolState(self, nlsPolarized, Ex, Ey, Phx, Phy):
        """Sets the default polarization state.

        These parameters correspond to the Polarization tab under
        the General settings.

        Parameters
        ----------
        nlsPolarized : integer
            if ``nlsPolarized > 0``, then default polarization state
            is unpolarized.
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
        polStateData : tuple
            the 5-tuple contains ``(nlsPolarized, Ex, Ey, Phax, Phay)``

        Notes
        -----
        The quantity ``Ex*Ex + Ey*Ey`` should have a value of 1.0
        although any values are accepted.

        See Also
        --------
        zGetPolState()
        """
        cmd = ("SetPolState,{:d},{:1.20g},{:1.20g},{:1.20g},{:1.20g}"
                .format(nlsPolarized,Ex,Ey,Phx,Phy))
        reply = self._sendDDEcommand(cmd)
        rs = reply.rsplit(",")
        polStateData = [int(float(elem)) if i==0 else float(elem)
                                       for i,elem in enumerate(rs[:-1])]
        return tuple(polStateData)

    def zSetSettingsData(self, number, data):
        """Sets the settings data used by a window in temporary storage
        before calling ``zMakeGraphicWindow()`` or ``zMakeTextWindow()``.
        The data may be retrieved using zGetSettingsData.

        Parameters
        ----------
        number : integer
            currently, only ``number = 0`` is supported. This number may be
            used to expand the feature in the future.
        data : ??

        Returns
        -------
        settingsData : string
            settings data returned by Zemax

        Notes
        -----
        Please refer to "How ZEMAX calls the client" in the Zemax manual.

        See Also
        --------
        zGetSettingsData()
        """
        cmd = "SettingsData,{:d},{}".format(number, data)
        reply = self._sendDDEcommand(cmd)
        return str(reply.rstrip())

    def zSetSolve(self, surfNum, code, *solveData):
        """Sets data for solves and/or pickups on the surface

        Parameters
        ----------
        surfNum : integer
            surface number for which the solve is to be set.
        code : integer
            surface parameter code for curvature, thickness, glass, conic,
            semi-diameter, etc. (refer to table surf_param_codes_for_setsolve_ 
            or use surface parameter mnemonic codes with signature 
            `ln.SOLVE_SPAR_XXX`. for e.g. `ln.SOLVE_SPAR_CURV`, etc. 
        solveData : splattered tuple
            the tuple of arguments are ``solvetype, param1, param2, param3, param4``.
            Refer to the table refer to table surf_param_codes_for_setsolve_ to
            construct the `solveData` sequence for specific solve type code.
            There are two ways of passing this parameter:
    
                1. As a tuple of the above arguments preceded by the ``*``
                   operator to flatten/splatter the tuple (see example below).
                2. As a sequence of arguments: 
                   ``solvetype, param1, param2, param3, param4`` or

            IMPORTANT:  
            (1) All parameters should be passed as there is no default 
                arguments to the function.
            (2) The order of parameters for `solveData` for code 5-17 
                do not match the order in the pop-up window for setting
                pickup solve in Zemax application. For others, i.e.
                when `solveData` is specified as `solvetype, param1, param2 ...`
                the order of `param1`, `param2` matches the corresponding
                order of parameters in Zemax solve window.
            (3) For `solvetypes` that has pickup column, use 0 for  
                "current column".

        Returns
        -------
        solveData : tuple
            tuple depending on the code value according to the table
            surf_param_codes_for_setsolve_ (same return as ``zGetSolve()``),
            if successful. The first element in the tuple is always the
            `solvetype`. -1 if the command failed.

        Notes
        -----
        1. The ``solvetype`` is an integer code, & the parameters have
           meanings that depend upon the solve type; see the chapter
           "SOLVES" in the Zemax manual for details. You may also use 
           the mnemonic codes with signature ln.SOLVE_XXX, such as 
           ln.SOLVE_CURV_FIXED, ln.SOLVE_CURV_FIXED, ln.SOLVE_THICK_VAR,
           etc. Additionally, it may also help to directly
           refer to the function body to quickly get an idea about the
           ``solvetype`` codes and parameters.
        2. If the ``solvetype`` is fixed, then the ``value`` in the
           ``solveData`` is ignored.
        3. Surface parameter codes
        
            .. _surf_param_codes_for_setsolve:
    
            ::
    
                Table : Surface parameter codes for zGetsolve() and zSetSolve()
    
                --------------------------------------------------------------------------
                   code           - Datum set/get by zGetSolve()/zSetSolve()
                --------------------------------------------------------------------------
                0 (curvature)     - solvetype, param1, param2, pickupcolumn
                1 (thickness)     - solvetype, param1, param2, param3, pickupcolumn
                2 (glass)         - solvetype (for solvetype = 0);
                                    solvetype, Index, Abbe, Dpgf (for solvetype = 1, model glass);
                                    solvetype, pickupsurf (for solvetype = 2, pickup);
                                    solvetype, index_offset, abbe_offset (for solvetype = 4, offset);
                                    solvetype (for solvetype=all other values)
                3 (semi-diameter) - solvetype, pickupsurf, pickupcolumn
                4 (conic)         - solvetype, pickupsurf, pickupcolumn
                5-16 (param 1-12) - solvetype, pickupsurf, offset, scalefactor, pickupcolumn
                17 (parameter 0)  - solvetype, pickupsurf, offset, scalefactor, pickupcolumn
                1001+ (extra      - solvetype, pickupsurf, scalefactor, offset, pickupcolumn
                data values 1+)     
    
                end-of-table
        4. If a parameter in the LDE is also present in the Multi-Configuration-Editor, 
           Zemax doesn't allow the solve on that parameter to be set in the LDE. Instead,
           change the "status" of that parameter to set a solve in the MCE using the 
           command `zSetMulticon()`.
        
        Examples
        --------
        To set a solve on the curvature (0) of surface number 6 such that
        the Marginal Ray angle (2) value is 0.1, use any of the following:

        >>> sdata = ln.zSetSolve(6, 0, *(2, 0.1))

        OR

        >>> sdata = ln.zSetSolve(6, 0, 2, 0.1 )
        
        OR
        
        >>> sdata = ln.zSetSolve(6, ln.SOLVE_SPAR_CURV, ln.SOLVE_CURV_MR_ANG, 0.1)

        See Also
        --------
        zSetMulticon() : for setting solves on parameters in Multi-Configuration-Editor; 
        zGetSolve(), zGetNSCSolve(), zSetNSCSolve(), zRemoveVariables().
        """
        if not solveData:
            print("Error [zSetSolve] No solve data passed.")
            return -1
        try:
            if code == self.SOLVE_SPAR_CURV:  # Solve specified on CURVATURE        
                
                if solveData[0] == self.SOLVE_CURV_FIXED:         
                    data = ''
                
                elif solveData[0] == self.SOLVE_CURV_VAR: # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_CURV_MR_ANG: # (M)
                    data = '{:1.20g}'.format(solveData[1]) # angle
                
                elif solveData[0] == self.SOLVE_CURV_CR_ANG: # (C)
                    data = '{:1.20g}'.format(solveData[1]) # angle
                
                elif solveData[0] == self.SOLVE_CURV_PICKUP:  # (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1], solveData[2], solveData[3])) # suface, scale-factor, column
                
                elif solveData[0] == self.SOLVE_CURV_MR_NORM: # (N)
                    data = ''
                
                elif solveData[0] == self.SOLVE_CURV_CR_NORM: # (N)
                    data = ''
                
                elif solveData[0] == self.SOLVE_CURV_APLAN: # (A)
                    data = ''
                
                elif solveData[0] == self.SOLVE_CURV_ELE_POWER: # (X)
                    data = '{:1.20g}'.format(solveData[1]) # power
                
                elif solveData[0] == self.SOLVE_CURV_CON_SURF: # (S)
                    data = '{:d}'.format(solveData[1]) # surface to be concentric to
                
                elif solveData[0] == self.SOLVE_CURV_CON_RADIUS: # (R)
                    data = '{:d}'.format(solveData[1]) # surface to be concentric with
                
                elif solveData[0] == self.SOLVE_CURV_FNUM: # (F)
                    data = '{:1.20g}'.format(solveData[1]) # paraxial f/#
                
                elif solveData[0] == self.SOLVE_CURV_ZPL: # (Z)
                    data = str(solveData[1])       # macro name
            
            elif code == self.SOLVE_SPAR_THICK:  # Solve specified on THICKNESS
                
                if solveData[0] == self.SOLVE_THICK_FIXED: 
                    data = ''
                
                elif solveData[0] == self.SOLVE_THICK_VAR:  # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_THICK_MR_HGT: # (M)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1], solveData[2]) # height, pupil zone
                
                elif solveData[0] == self.SOLVE_THICK_CR_HGT: # (C)
                    data = '{:1.20g}'.format(solveData[1])   # height
                
                elif solveData[0] == self.SOLVE_THICK_EDGE_THICK: # (E)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1], solveData[2]) # thickness, radial height (0 for semi-diameter)
                
                elif solveData[0] == self.SOLVE_THICK_PICKUP: # (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1], solveData[2], solveData[3], solveData[4])) # surface, scale-factor, offset, column
                
                elif solveData[0] == self.SOLVE_THICK_OPD: # (O)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1], solveData[2]) # opd, pupil zone
                
                elif solveData[0] == self.SOLVE_THICK_POS: # (T)
                    data = '{:d},{:1.20g}'.format(solveData[1], solveData[2]) # surface, length from surface
                
                elif solveData[0] == self.SOLVE_THICK_COMPENSATE: # (S)
                    data = '{:d},{:1.20g}'.format(solveData[1], solveData[2]) # surface, sum of surface thickness
                
                elif solveData[0] == self.SOLVE_THICK_CNTR_CURV: # (X)
                    data = '{:d}'.format(solveData[1]) # surface to be at the COC of
                
                elif solveData[0] == self.SOLVE_THICK_PUPIL_POS: # (U)
                    data = ''
                
                elif solveData[0] == self.SOLVE_THICK_ZPL: # (Z)
                    data = str(solveData[1])       # macro name
            
            elif code == self.SOLVE_SPAR_GLASS: # GLASS
                
                if solveData[0] == self.SOLVE_GLASS_FIXED:
                    data = ''
                
                elif solveData[0] == self.SOLVE_GLASS_MODEL: 
                    data = ('{:1.20g},{:1.20g},{:1.20g}'
                    .format(solveData[1], solveData[2], solveData[3])) # index Nd, Abbe Vd, Dpgf
                
                elif solveData[0] == self.SOLVE_GLASS_PICKUP: # (P)
                    data = '{:d}'.format(solveData[1]) # surface
                
                elif solveData[0] == self.SOLVE_GLASS_SUBS: # (S)
                    data = str(solveData[1])      # catalog name
                
                elif solveData[0] == self.SOLVE_GLASS_OFFSET: # (O)
                    data = '{:1.20g},{:1.20g}'.format(solveData[1], solveData[2]) # index Nd offset, Abbe Vd offset
            
            elif code == self.SOLVE_SPAR_SEMIDIA:   # Solve specified on SEMI-DIAMETER
                
                if solveData[0] == self.SOLVE_SEMIDIA_AUTO:
                    data = ''
                
                elif solveData[0] == self.SOLVE_SEMIDIA_FIXED: # (U)
                    data = ''
                
                elif solveData[0] == self.SOLVE_SEMIDIA_PICKUP:  # (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1], solveData[2], solveData[3])) # surface, scale-factor, column
                
                elif solveData[0] == self.SOLVE_SEMIDIA_MAX: # (M)
                    data = ''
                
                elif solveData[0] == self.SOLVE_SEMIDIA_ZPL:  # (Z)
                    data = str(solveData[1])       # macro name
            
            elif code == self.SOLVE_SPAR_CONIC:  # Solve specified on CONIC
                
                if solveData[0] == self.SOLVE_CONIC_FIXED:  
                    data = ''
                
                elif solveData[0] == self.SOLVE_CONIC_VAR: # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_CONIC_PICKUP: # (P)
                    data = ('{:d},{:1.20g},{:d}'
                    .format(solveData[1], solveData[2], solveData[3])) # surface, scale-factor, column
                
                elif solveData[0] == self.SOLVE_CONIC_ZPL: # (Z)
                    data = str(solveData[1])       # macro name
            
            elif self.SOLVE_SPAR_PAR1 <= code <= self.SOLVE_SPAR_PAR12:  # Solve specified on PARAMETERS 1-12  
                
                if solveData[0] == self.SOLVE_PARn_FIXED:  
                    data = ''
                
                elif solveData[0] == self.SOLVE_PARn_VAR:  # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_PARn_PICKUP: # (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1],solveData[2],solveData[3],solveData[4])) # surface, scale-factor, offset, column
                     
                elif solveData[0] == self.SOLVE_PARn_CR:  # (C)
                    data = '{:d},{:1.20g}'.format(solveData[1], solveData[2]) # field, wavelength
                
                elif solveData[0] == self.SOLVE_PARn_ZPL: # (Z)
                    data = str(solveData[1]) # macro name
            
            elif code == self.SOLVE_SPAR_PAR0:  # Solve specified on PARAMETER 0
                
                if solveData[0] == self.SOLVE_PAR0_FIXED:  
                    data = ''
                
                elif solveData[0] == self.SOLVE_PAR0_VAR: # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_PAR0_PICKUP: # (P)
                    data = '{:d}'.format(solveData[1]) # surface
            
            elif code > 1000:  # Solve specified on EXTRA DATA VALUES
                
                if solveData[0] == self.SOLVE_EDATA_FIXED: 
                    data = ''
                
                elif solveData[0] == self.SOLVE_EDATA_VAR: # (V)
                    data = ''
                
                elif solveData[0] == self.SOLVE_EDATA_PICKUP: # (P)
                    data = ('{:d},{:1.20g},{:1.20g},{:d}'
                    .format(solveData[1], solveData[2], solveData[3], solveData[4])) # surface, scale-factor, offset, column
                
                elif solveData[0] == self.SOLVE_EDATA_ZPL: # (Z)
                    data = str(solveData[1]) # macro name
        
        except IndexError:
            print("Error [zSetSolve]: Check number of solve parameters!")
            return -1
        #synthesize the command to pass to zemax
        if data:
            cmd = ("SetSolve,{:d},{:d},{:d},{}"
                  .format(surfNum, code, solveData[0], data))
        else:
            cmd = ("SetSolve,{:d},{:d},{:d}"
                  .format(surfNum, code, solveData[0]))
        reply = self._sendDDEcommand(cmd)
        solveData = _process_get_set_Solve(reply)
        return solveData

    def zSetSurfaceData(self, surfNum, code, value, arg2=None):
        """Sets surface data on a sequential lens surface

        Parameters
        ----------
        surfNum : integer
            the surface number
        code : integer
            number (Refer to the table surf_data_codes_ in the docstring
            of ``zGetSurfaceData()``). You may also use the surface data 
            mnemonic codes with signature ln.SDAT_XXX, e.g. ln.SDAT_TYPE, 
            ln.SDAT_CURV, ln.SDAT_THICK, etc 
        value : string or float
            string if ``code`` is 0, 1, 4, 7 or 9,  else float
        arg2 : optional
            required for item codes above 70.

        Returns
        -------
        surface_data : string or numeric
            the returned data depends on the ``code``. Refer to the table
            surf_data_codes_ (in docstring of ``zGetSurfaceData()``) for
            details

        See Also
        --------
        zGetSurfaceData()
        """
        cmd = "SetSurfaceData,{:d},{:d}".format(surfNum,code)
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

    def zSetSurfaceParameter(self, surfNum, param, value):
        """Set surface parameter data.

        Parameters
        ----------
        surfNum : integer
            surface number of the surface
        param : integer
            parameter (Par in LDE) number being set
        value : float
            value to set for the ``param``

        Returns
        -------
        paramData : float
            the parameter value

        See Also
        --------
        zSetSurfaceData(), zGetSurfaceParameter()
        """
        cmd = ("SetSurfaceParameter,{:d},{:d},{:1.20g}"
               .format(surfNum,param,value))
        reply = self._sendDDEcommand(cmd)
        return float(reply)

    def zSetSystem(self, unitCode=0, stopSurf=1, rayAimingType=0, useEnvData=0,
                   temp=20.0, pressure=1, globalRefSurf=1):
        """Sets a number of general systems property (General Lens Data)

        Parameters
        ----------
        unitCode : integer, optional
            lens units code (0, 1, 2, or 3 for mm, cm, in, or M)
        stopSurf : integer, optional
            the stop surface number
        rayAimingType : integer, optional
            ray aiming type (0, 1, or 2 for off, paraxial or real)
        useEnvData : integer, optional
            use environment data flag (0 or 1 for no or yes) [ignored]
        temp : float, optional
            the current temperature
        pressure : float, optional
            the current pressure
        globalRefSurf : integer, optional
            the global coordinate reference surface number

        Returns
        -------
        numSurfs : integer
            number of surfaces
        unitCode : integer
            lens units code (0, 1, 2, or 3 for mm, cm, in, or M)
        stopSurf : integer
            the stop surface number
        nonAxialFlag : integer
            flag to indicate if system is non-axial symmetric (0 for axial,
            1 if not axial);
        rayAimingType : integer
            ray aiming type (0, 1, or 2 for off, paraxial or real)
        adjustIndex : integer
            adjust index data to environment flag (0 if false, 1 if true)
        temp : float
            the current temperature
        pressure : float
            the current pressure
        globalRefSurf : integer
            the global coordinate reference surface number

        See Also
        --------
        zSetSystemAper():
            for setting the system aperture such as aperture type, aperture
            value, etc.
        zGetSystem(), zGetSystemAper(), zGetAperture(), zSetAperture()
        """
        cmd = ("SetSystem,{:d},{:d},{:d},{:d},{:1.20g},{:1.20g},{:d}"
              .format(unitCode,stopSurf,rayAimingType,useEnvData,temp,pressure,
               globalRefSurf))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        systemData = tuple([float(elem) if (i==6) else int(float(elem))
                                                  for i,elem in enumerate(rs)])
        return systemData

    def zSetSystemAper(self, aType, stopSurf, aperVal):
        """Sets the lens system aperture and corresponding data.

        Parameters
        ----------
        aType : integer indicating the system aperture
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
        aType : integer
            indicating the system aperture as follows:

            | 0 = entrance pupil diameter (EPD)
            | 1 = image space F/#         (IF/#)
            | 2 = object space NA         (ONA)
            | 3 = float by stop           (FBS)
            | 4 = paraxial working F/#    (PWF/#)
            | 5 = object cone angle       (OCA)

        stopSurf : integer
            stop surface
        value : float
            if aperture type is "float by stop" value is stop surface
            semi-diameter else value is the sytem aperture

        Notes
        -----
        The returned tuple is the same as the returned tuple of
        ``zGetSystemAper()``

        See Also
        --------
        zGetSystem(), zGetSystemAper()
        """
        cmd = ("SetSystemAper,{:d},{:d},{:1.20g}".format(aType, stopSurf, aperVal))
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        systemAperData = tuple([float(elem) if i==2 else int(float(elem))
                                for i, elem in enumerate(rs)])
        return systemAperData

    def zSetSystemProperty(self, code, value1, value2=0):
        """Sets system properties

        Parameters
        ----------
        code : integer
            value that defines the specific system property to be set
            (see the table system_property_codes_ in docstring of
            ``zGetSystemProperty()``)
        value1 : integer or float or string
            the nature and type of ``value1`` depends on the ``code``
        value2 : integer or float, oprional
            the nature and type of ``value2`` depends on the ``code``.
            Ignored if not used

        Returns
        -------
        sysPropData : string or numeric
            system property data (refer to the table system_property_codes_
            in docstring of ``zGetSystemProperty()``)

        See Also
        --------
        zGetSystemProperty()
        """
        cmd = ("SetSystemProperty,{c:d},{v1},{v2}".format(c=code, v1=value1, v2=value2))
        reply = self._sendDDEcommand(cmd)
        sysPropData = _process_get_set_SystemProperty(code,reply)
        return sysPropData

    def zSetTol(self, operNum, col, value):
        """Sets the tolerance operand data.

        Parameters
        ----------
        operNum : integer
            tolerance operand number (row number in the tolerance editor,
            when greater than 0)
        col : integer
            * 1 for tolerance Type;
            * 2-4 for int1 - int3;
            * 5 for min;
            * 6 for max;
        value : string or float
            4-character string (tolerancing operand code) if ``col==1``,
            else float value to set

        Returns
        -------
        toleranceData : number or tuple or errorCode
            the ``toleranceData`` is a number or a 6-tuple, depending
            upon ``operNum`` as follows:

            * if ``operNum = 0``, then ``toleranceData`` is a number
              indicating the number of tolerance operands defined.
            * if ``operNum > 0``, then ``toleranceData`` is a tuple
              with elements ``(tolType, int1, int2, min, max, int3)``
            * Returns -1 if an error occurs.

        See Also
        --------
        zSetTolRow(), zGetTol(),
        """
        if col == 1: # value is string code for the operand
            if zo.isZOperand(str(value),2):
                cmd = "SetTol,{:d},{:d},{}".format(operNum,col,value)
            else:
                return -1
        else:
            cmd = "SetTol,{:d},{:d},{:1.20g}".format(operNum,col,value)
        reply = self._sendDDEcommand(cmd)
        if operNum == 0: # returns just the number of operands
            return int(float(reply.rstrip()))
        else:
            return _process_get_set_Tol(operNum,reply)
        # FIX !!! currently, I am not able to set more than 1 row in the tolerance
        # editor, through this command. I don't find anything like zInsertTol ...
        # A similar function exist for Multi-Configuration editor (zInsertMCO) and
        # for Multi-function editor (zInsertMFO). May need to contact Zemax Support.

    def zSetUDOItem(self, bufferCode, dataNum, data):
        """This function is used to pass just one datum computed by the
        client program to the Zemax optimizer.

        Parameters
        ----------
        bufferCode : integer
            the integer value provided by Zemax to the client that
            uniquely identifies the correct lens.
        dataNum : integer
            ?
        data : float
            data item number being passed

        Returns
        -------
          ?

        Notes
        -----
        1. The only time this item name should be used is when implementing
           a User Defined Operand, or UDO. UDO's are described in
           "Optimizing with externally compiled programs" in the Zemax
           manual.

        2. After the last data item has been sent, the buffer must be
           closed using the ``zCloseUDOData()`` function before the
           optimization may proceed. A typical implementation may consist
           of the following series of function calls:

           ::

            ln.zSetUDOItem(bufferCode, 0, value0)
            ln.zSetUDOItem(bufferCode, 1, value1)
            ln.zSetUDOItem(bufferCode, 2, value2)
            ln.zCloseUDOData(bufferCode)

        See Also
        --------
        zGetUDOSystem(), zCloseUDOData().
        """
        cmd = "SetUDOItem,{:d},{:d},{:1.20g}".format(bufferCode, dataNum, data)
        reply = self._sendDDEcommand(cmd)
        return _regressLiteralType(reply.rstrip())
        # FIX !!! At this time, I am not sure what is the expected return.

    def zSetVig(self):
        """Request Zemax to set the vignetting factors automatically.

        Parameters
        ----------
        None

        Returns
        -------
        retVal : integer
            0 = success, -1 = fail

        Notes
        -----
        Calling this function is equivalent to clicking the "Set Vig"
        button from the "Field Data" window. For more information on
        how Zemax calculates the vignetting factors automatically, please
        refer to "Vignetting factors" under the "Systems Menu" chapter in
        the Zemax Manual.
        """
        retVal = -1
        reply = self._sendDDEcommand("SetVig")
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zSetWave(self, n, arg1, arg2):
        """Sets wavelength data in the Zemax DDE server

        There are 2 ways to use this function:

            ``zSetWave(0, primary, number) -> waveData``

             OR

            ``zSetWave(n, wavelength, weight) -> waveData``

        Parameters
        ----------
        [Case: ``n=0``]:

        n : 0
            the function sets general wavelength data
        arg1 : integer
            primary wavelength number to set
        arg2 : integer
            total number of wavelengths to set

        [Case: ``0 < n <= number-of-wavelengths``]:

        n : integer (> 0)
            wavelength number to set
        arg1 : float
            wavelength in micrometers
        arg2 : float
            weight

        Returns
        -------
        The function returns a tuple. The elements in the tuple has
        different meaning depending on the value of ``n``.

        [Case: ``n=0``]:

        primary : integer
            number indicating the primary wavelength
        number : integer
            number of wavelengths currently defined

        [Case: ``0 < n <= number-of-wavelengths``]:

        wavelength : float
            value of the specific wavelength
        weight : float
            weight of the specific wavelength

        Notes
        -----
        The returned tuple is exactly same in structure and contents to
        that returned by ``zGetWave()``.

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

    def zWindowMaximize(self, windowNum=0):
        """Maximize the main Zemax window or any analysis window Zemax
        currently displayed.

        Parameters
        ----------
        windowNum : integer
            the window number. use 0 for the main Zemax window

        Returns
        -------
        retVal : integer
            0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowMaximize,{:d}".format(windowNum))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowMinimize(self, windowNum=0):
        """Minimize the main Zemax window or any analysis window Zemax
        currently

        Parameters
        -----------
        windowNum : integer
            the window number. use 0 for the main Zemax window

        Returns
        -------
        retVal : integer
            0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowMinimize,{:d}".format(windowNum))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    def zWindowRestore(self, windowNum=0):
        """Restore the main Zemax window or any analysis window to it's
        previous size and position.

        Parameters
        ----------
        windowNum : integer
            the window number. use 0 for the main Zemax window

        Returns
        -------
        retVal : integer
            0 if success, -1 if failed.
        """
        retVal = -1
        reply = self._sendDDEcommand("WindowRestore,{:d}".format(windowNum))
        if 'OK' in reply.split():
            retVal = 0
        return retVal

    #%% ADDITIONAL FUNCTIONS 

    # -------------------------------------------------------
    # Editor function for both getting and setting parameters
    # -------------------------------------------------------
    def zGetOperandRow(self, row):
        """Returns a row of the Multi Function Editor

        Parameters
        ----------
        row : integer
            the operand row number

        Returns
        -------
        opertype : string
            operand type, column 1 in MFE
        int1 : integer or string 
            column 2 in MFE. The column 2 is a string, usually when opertype 
            is 'BLNK', and there is some comments in the second column 
        int2 : integer
            column 3 in MFE
        data1 : float
            column 4 in MFE
        data2 : float
            column 5 in MFE
        data3 : float
            column 6 in MFE
        data4 : float
            column 7 in MFE
        data5 : float
            column 12 in MFE
        data6 : float
            column 13 in MFE
        tgt : float
            target
        wgt : float
            weight
        value : float
            value
        percentage : float
            percentage contribution

        See Also
        --------
        zGetOperand(), zSetOperandRow()
        """
        operData = []
        for i in range(1,8):
            operData.append(self.zGetOperand(row=row, column=i))
        for i in range(12, 14):
            operData.append(self.zGetOperand(row=row, column=i))
        for i in range(8, 12):
            operData.append(self.zGetOperand(row=row, column=i))
        rowdat = _co.namedtuple('OperandData', ['opertype', 'int1', 'int2', 'data1',
                                'data2', 'data3', 'data4', 'data5', 'data6', 'tgt',
                                'wgt', 'value', 'percentage'])
        return rowdat._make(operData)

    def zSetOperandRow(self, row, opertype, int1=None, int2=None, data1=None, data2=None,
                     data3=None, data4=None, data5=None, data6=None, tgt=None, wgt=None):
        """Sets a row in the Merit Function Editor

        Parameters
        ----------
        row : integer
            operand row number in the MFE
        opertype : string
            operand type
        int1 : integer, optional
            column 2 in MFE
        int2 : integer, optional
            column 3 in MFE
        data1 : float, optional
            column 4 in MFE
        data2 : float, optional
            column 5 in MFE
        data3 : float, optional
            column 6 in MFE
        data4 : float, optional
            column 7 in MFE
        data5 : float, optional
            column 12 in MFE
        data6 : float, optional
            column 13 in MFE
        tgt : float, optional
            target
        wgt : float, optional
            weight

        Returns
        -------
        the contents of the row. (same as that returned by
        ``zGetOperandRow()``)

        Notes
        -----
        1. Use ``zInsertMFO()`` to insert a new row in the MFE at a
           specified row number.
        2. To update the merit function after calling ``zSetOperand()``,
           call ``zOptimize()`` with the number of cycles set to -1.

        See Also
        --------
        zInsertMFO(), zSetOperand(), zOperandValue(), zGetOperand()
        """
        values1_9 = (opertype, int1, int2, data1, data2, data3, data4, tgt, wgt)
        values12_13 = (data5, data6)
        for i, val in enumerate(values1_9):
            if val is not None:
                self.zSetOperand(row=row, column=i+1, value=val)
        for i, val in enumerate(values12_13):
            if val is not None:
                self.zSetOperand(row=row, column=i+12, value=val)
        return self.zGetOperandRow(row)

    # -------------------
    # System functions
    # -------------------
    def zGetAngularMagnification(self, wave=None):
        """Get angular magnification of paraxial system.

        The angular magnification is defined as the ratio of the image space 
        paraxial chief ray angle to the object space paraxial chief ray angle 

        Parameters
        ---------- 
        wave : integer, optional 
            the wavelength defined by `wave`. If `None`, the primary wave 
            number is used.

        Returns
        ------- 
        amag : real 
            angular magnification, at least one non-zero field point is 
            defined in the Field Data Editor.
            Returns error code -999 if only on-axis field is defined. 
            See Notes.

        Notes
        ----- 
        Zemax returns zero (0) for angular magnification if the only field 
        defined in the field editor is the on-axis field.

        See Also
        -------- 
        zGetPupilMagnification()    
        """
        if wave==None:
            wave = self.zGetPrimaryWave()
        if self.zAnyOffAxisField():
            return self.zOperandValue('AMAG', wave)
        else:
            return -999

    def zGetMagnification(self):
        """Returns the real magnification evaluated as the ratio of the image 
        height to the object height 

        Parameters
        ---------- 
        None 

        Returns
        ------- 
        mag : real 
            real magnification. see Notes.

        Notes
        ----- 
        1. The function returns the real magnification of the system. It is 
           affected by distortions, and the actual location of the image 
           surface. For paraxial magnification use `zGetFirst().paraMag`
        """
        objHt = self.zGetSemiDiameter(0)
        if objHt:
            rtd = self.zGetTrace(waveNum=1, mode=0, surf=-1, hx=0, hy=1, px=0, py=0)        
            return rtd.y/objHt
        else:
            return 0.0

    def zGetNumField(self):
        """Returns the total number of fields defined

        Equivalent to ZPL macro ``NFLD``

        Parameters
        ----------
        None

        Returns
        -------
        nfld : integer
            number of fields defined
        """
        return self.zGetSystemProperty(101)

    def zAnyOffAxisField(self):
        """Returns `True` if at least one off-axis X-Field or Y-Field is 
        defined in the Field Data Editor.

        Fields with zero weights are also considered to be "defined".

        Parameters
        ---------- 
        None 

        Returns
        ------- 
        retVal : bool 
            `True` if any off-axis field is found, else `False`
        """
        fdata = self.zGetFieldTuple()
        fx = [f[0] for f in fdata]
        fy = [f[1] for f in fdata]
        return any(fy) or any(fx)

    def zGetFieldTuple(self):
        """Get all field data in a single n-tuple.

        Parameters
        ----------
        None

        Returns
        -------
        fieldDataTuple: n-tuple (``0 < n <= 12``)
            the tuple elements represent field loactions with each element
            containing all 8 field parameters.

        Examples
        --------
        This example shows the namedtuple returned by ``zGetFieldTuple()``

        >>> ln.zGetFieldTuple()
        (fieldData(xf=0.0, yf=0.0, wgt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0),
         fieldData(xf=0.0, yf=14.0, wgt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0),
         fieldData(xf=0.0, yf=20.0, wgt=1.0, vdx=0.0, vdy=0.0, vcx=0.0, vcy=0.0, van=0.0))

        See Also
        --------
        zGetField(), zSetField(), zSetFieldTuple()
        """
        fieldCount = self.zGetField(0)[1]
        fd = _co.namedtuple('fieldData', ['xf', 'yf', 'wgt',
                                          'vdx', 'vdy',
                                          'vcx', 'vcy', 'van'])
        fieldData = []
        for i in range(fieldCount):
            reply = self._sendDDEcommand('GetField,' + str(i+1))
            rs = reply.split(',')
            data = fd._make([float(elem) for elem in rs])
            fieldData.append(data)
        return tuple(fieldData)

    def zSetFieldTuple(self, ftype, norm, fields):
        """Sets all field points from a 2D field tuple

        Parameters
        ----------
        ftype : integer
            the field type (0 = angle, 1 = object height, 2 = paraxial
            image height, and 3 = real image height)
        norm : integer 0 or 1
            the field normalization (0=radial, 1=rectangular)
        fields : n-tuple
            the input field data tuple is an N-D tuple (0 < N <= 12) with
            every dimension representing a single field location. It can
            be constructed as shown in the example (see below)

        Returns
        -------
        fields : n-tuple
            the output field data tuple is also a N-D tuple similar to the
            ``fields``, except that for each field location all
            8 field parameters are returned.

        Examples
        --------
        The following example sets 3 field points defined as angles with
        field normalization = 1:

            * xf=0.0, yf=0.0, wgt=1.0, vdx=vdy=vcx=vcy=van=0.0
            * xf=0.0, yf=5.0, wgt=1.0
            * xf=0.0, yf=10.0

        >>> ln.zSetFieldTuple(0, 0,
                              (0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0),
                              (0.0, 5.0, 1.0),
                              (0.0, 7.0))

        See Also
        --------
        zSetField(), zGetField(), zGetFieldTuple()
        """
        fieldCount = len(fields)
        if not 0 < fieldCount <= 12:
            raise ValueError('Invalid number of fields')
        cmd = ("SetField,{:d},{:d},{:d},{:d}"
              .format(0, ftype, fieldCount, norm))
        self._sendDDEcommand(cmd)
        retFields = []
        for i in range(fieldCount):
            fieldData = self.zSetField(i+1,*fields[i])
            retFields.append(fieldData)
        return tuple(retFields)

    def zGetNumSurf(self):
        """Return the total number of surfaces defined

        Equivalent to ZPL macro ``NSUR``

        Parameters
        ----------
        None

        Returns
        -------
        nsur : integer
            number of surfaces defined

        Notes
        -----
        The count doesn't include the object (OBJ) surface.
        """
        return self.zGetSystem()[0]

    def zGetNumWave(self):
        """Return the total number of wavelengths defined

        Equivalent to ZPL macro ``NWAV``

        Parameters
        ----------
        None

        Returns
        -------
        nwav : integer
            number of wavelengths defined
        """
        return self.zGetSystemProperty(201)

    def zGetPrimaryWave(self):
        """Return the primary wavelength number

        Equivalent to ZPL macro ``PWAV``

        Parameters
        ----------
        None

        Returns
        -------
        primary_wave_number : integer
            primary wavelength number

        Notes
        -----
        To get the primary wavelength (in microns) do the following:

        >>> ln.zGetWave(ln.zGetPrimaryWave()).wavelength

        See Also
        --------
        zSetPrimaryWave()
        """
        return self.zGetSystemProperty(200)

    def zSetPrimaryWave(self, waveNum):
        """Sets the primary wavelength for the lens in Zemax DDE server.

        Parameters
        ----------
        waveNum : integer
            the wavelength number to set as primary

        Returns
        -------
        primary : integer
            number indicating the primary wavelength
        number : integer
            number of wavelengths currently defined

        See Also
        --------
        zGetPrimaryWave(), zSetWave(), zSetWave(), zSetWaveTuple(),
        zGetWaveTuple()
        """
        waveData = self.zGetWave(0)
        cmd = "SetWave,{:d},{:d},{:d}".format(0, waveNum, waveData[1])
        reply = self._sendDDEcommand(cmd)
        rs = reply.split(',')
        waveData = tuple([int(elem) for elem in rs])
        return waveData

    def zGetWaveTuple(self):
        """Gets data on all defined wavelengths

        Parameters
        ----------
        None

        Returns
        -------
        waveDataTuple : 2D tuple
            the first dimension (first subtuple) contains the wavelengths
            and the second dimension containing the weights as follows:
            ((wave1, wave2, wave3 ,..., waveN),(wgt1, wgt2, wgt3,..., wgtN))

        Notes
        -----
        This function is similar to "zGetWaveDataMatrix()" in MZDDE.

        Examples
        --------
        This example shows the named tuple returned by the function

        >>> ln.zGetWaveTuple()
        waveDataTuple(wavelengths=(0.48, 0.55, 0.65), weights=(0.800000012, 1.0, 0.800000012))

        See Also
        --------
        zSetWaveTuple(), zGetWave(), zSetWave()
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

    def zSetWaveTuple(self, waves):
        """Sets wavelength and weight data from a matrix.

        Parameters
        ----------
        waves : 2-D tuple
            the input wave data tuple is a 2D tuple with the first
            dimension (first sub-tuple) containing the wavelengths
            and the second dimension containing the weights as:

            ``((wave1,wave2,wave3,...,waveN),(wgt1,wgt2,wgt3,...,wgtN))``

            The first wavelength (wave1) is assigned to be the primary
            wavelength. To change the primary wavelength use
            ``zSetWavePrimary()``

        Returns
        -------
        retWaves : 2-D tuple
            the output wave data tuple is also a 2D tuple similar to the
            ``waves``.

        See Also
        --------
        zGetWaveTuple(), zSetWave(), zSetWavePrimary()
        """
        waveCount = len(waves[0])
        retWaves = [[],[]]
        self.zSetWave(0,1,waveCount) # Set no. of wavelen & the wavelen to 1
        for i in range(waveCount):
            cmd = ("SetWave,{:d},{:1.20g},{:1.20g}".format(i+1,waves[0][i],waves[1][i]))
            reply = self._sendDDEcommand(cmd)
            rs = reply.split(',')
            retWaves[0].append(float(rs[0])) # store the wavelength
            retWaves[1].append(float(rs[1])) # store the weight
        return (tuple(retWaves[0]),tuple(retWaves[1]))

    def zSetNSCPositionTuple(self, surfNum, objNum, x=0.0, y=0.0, z=0.0, 
                             tiltX=0.0, tiltY=0.0, tiltZ=0.0, material=''):
        """Sets position and tilt data for NSC objects

        Parameters
        ----------
        surfNum : integer
            the surface number. Use 1 if Non-Sequential program mode
        objNum : integer
            the object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        material : string, optional
            valid string code to specify the material

        Returns
        -------
        nscPosData : tuple
            a 7-tuple containing x, y, z, tilt-x, tilt-y, tilt-z, material

        See Also
        --------
        zSetNSCPositionTuple(), zGetNSCPosition()
        """
        for code, item in enumerate((x, y, z, tiltX, tiltY, tiltZ, material), 1):
            self.zSetNSCPosition(surfNum, objNum, code, item)
        return self.zGetNSCPosition(surfNum, objNum)

    def zSetTolRow(self, operNum, tolType, int1, int2, int3, minT, maxT):
        """Helper function to set all the elements of a row (given
        by ``operNum``) in the tolerance editor.

        Parameters
        ----------
        operNum : integer (> 0)
            the tolerance operand number (row number in the tolerance
            editor)
        tolType : string
            4-character string (tolerancing operand code)
        int1 : integer
            'int1' parameter
        int2 : integer
            'int2' parameter
        int3 : integer
            'int3' parameter
        minT : float
            minimum value
        maxT : float
            maximum value

        Returns
        -------
        tolData : tolerance data or errorCode
            the data for the row indicated by the ``operNum``
            if successful, else -1
        """
        tolData = self.zSetTol(operNum, 1, tolType)
        if tolData != -1:
            self.zSetTol(operNum, 2, int1)
            self.zSetTol(operNum, 3, int2)
            self.zSetTol(operNum, 4, int3)
            self.zSetTol(operNum, 5, minT)
            self.zSetTol(operNum, 6, maxT)
            return self.zGetTol(operNum)
        else:
            return -1

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
        
    # -------------------
    # Analysis functions
    # -------------------
    
    # Spot diagram analysis functions
    def zSpiralSpot(self, hx, hy, waveNum, spirals, rays, mode=0):
        """Returns positions and intensity of rays traced in a spiral
        over the entrance pupil to the image surface.

        The final destination of the rays is the image surface.

        Parameters
        ----------
        hx : float
            normalized field height along x axis
        hy : float
            normalized field height along y axis
        waveNum : integer
            wavelength number as in the wavelength data editor
        spirals : integer 
            number of spirals 
        rays : integer 
            total number of rays to trace 
        mode : integer (0 or 1)
            0 = real; 1 = paraxial ray trace

        Returns
        -------
        rayInfo : 4-tuple
            (x, y, z, intensity)

        Notes
        -----
        This function imitates its namesake from MZDDE toolbox.
        Unlike the ``spiralSpot()`` of MZDDE, there is no need to call
        ``zLoadLens()`` before calling ``zSpiralSpot()``.
        """
        # Calculate the ray pattern on the pupil plane
        pi, cos, sin = _math.pi, _math.cos, _math.sin
        lastAng = spirals*2*pi
        delta_t = lastAng/(rays - 1) 
        theta = lambda dt, rays: (i*dt for i in range(rays))
        r = (i/(rays-1) for i in range(rays))
        pXY = ((r*cos(t), r*sin(t)) for r, t in _izip(r, theta(delta_t, rays)))
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

    # POP analysis functions
    def zGetPOP(self, settingsFile=None, displayData=False, txtFile=None,
                keepFile=False, timeout=None):
        """Returns Physical Optics Propagation (POP) data

        Parameters
        ----------
        settingsFile : string, optional
            * if passed, the POP will be called with this configuration
              file;
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens file post-fixed with
              "_pyzdde_POP.CFG" is present, the settings from this file
              will be used;
            * if no ``settingsFile`` and no file name post-fixed with
              "_pyzdde_POP.CFG" is found, but a config file with the same
              name as the lens file is present, the settings from that
              file will be used;
            * if no settings file is found, then a default settings will
              be used
        displayData : bool
            if ``true`` the function returns the 2D display data; default
            is ``false``
        txtFile : string, optional
            if passed, the POP data file will be named such. Pass a
            specific ``txtFile`` if you want to dump the file into a
            separate directory.
        keepFile : bool, optional
            if ``False`` (default), the POP text file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the POP text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds

        Returns
        -------
        popData : tuple
            popData is a 1-tuple containing just ``popInfo`` (see below)
            if ``displayData`` is ``False`` (default).
            If ``displayData`` is ``True``, ``popData`` is a 2-tuple
            containing ``popInfo`` (a tuple) and ``powerGrid`` (a 2D list):

            popInfo : named tuple
                surf : integer
                    surface number at which the POP is analysis was done
                peakIrr/ cenPhase : float
                    the peak irradiance is the maximum power per unit area
                    at any point in the beam, measured in source units per
                    lens unit squared. It returns center phase if the data
                    type is "Phase" in POP settings
                totPow : float
                    the total power, or the integral of the irradiance
                    over the entire beam if data type is "Irradiance" in
                    POP settings. This field is blank for "Phase" data
                fibEffSys : float
                    the efficiency of power transfer through the system
                fibEffRec : float
                    the efficiency of the receiving fiber
                coupling : float
                    the total coupling efficiency, the product of the
                    system and receiver efficiencies
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
                a two-dimensional list of the powers in the analysis grid
                if ``displayData`` is ``true``

        Notes
        -----
        The function returns ``None`` for any field which was not
        found in POP text file. This is most common in the case of
        ``fiberEfficiency_system`` and ``fiberEfficiency_receiver``
        as they need to be set explicitly in the POP settings

        See Also
        --------
        zSetPOPSettings(), zModifyPOPSettings()
        """
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, 'Pop')
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, 'Pop', cfgFile, getTextFlag,
                                timeout)
        assert ret == 0
        # get line list
        line_list = _readLinesFromFile(_openFile(textFileName))

        # Get data type ... phase or Irradiance?
        find_irr_data = _getFirstLineOfInterest(line_list, 'POP Irradiance Data',
                                                patAtStart=False)
        data_is_irr = False if find_irr_data is None else True
        # Get the Surface number and Grid size
        grid_line_num = _getFirstLineOfInterest(line_list, 'Grid size')
        surf_line = line_list[grid_line_num - 1]
        surf = int(_re.findall(r'\d{1,4}', surf_line)[0]) # assume: first int num in the line 
                                 # is surf number. surf comment can have int or float nums 
        grid_line = line_list[grid_line_num]
        grid_x, grid_y = [int(i) for i in _re.findall(r'\d{2,5}', grid_line)]

        # Point spacing
        pts_line = line_list[_getFirstLineOfInterest(line_list, 'Point spacing')]
        pat = r'-?\d\.\d{4,6}[Ee][-\+]\d{2,3}'
        pts_x, pts_y =  [float(i) for i in _re.findall(pat, pts_line)]

        width_x = pts_x*grid_x
        width_y = pts_y*grid_y
        
        if data_is_irr:
            # Peak Irradiance and Total Power
            pat_i = r'-?\d\.\d{4,6}[Ee][-\+]\d{2,3}' # pattern for P. Irr, T. Pow,
            peakIrr, totPow = None, None
            pi_tp_line = _getFirstLineOfInterest(line_list, 'Peak Irradiance') 
            if pi_tp_line: # Transfer magnitude doesn't have Peak Irradiance info
                pi_tp_line = line_list[pi_tp_line]
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
            #cp_line = line_list[_getFirstLineOfInterest(line_list, 'Center Phase')]
            cp_line = _getFirstLineOfInterest(line_list, 'Center Phase')
            if cp_line: # Transfer magnitude / Phase doesn't have Center Phase info
                cp_line = line_list[cp_line]
                cp = _re.search(pat_p, cp_line)
                if cp:
                    centerPhase = float(cp.group())
        # Pilot_size, Pilot_Waist, Pos, Rayleigh [... available for
        # both Phase and Irr data]
        pat_fe = r'\d\.\d{6}'   # pattern for fiber efficiency
        pat_pi = r'-?\d\.\d{4,6}[Ee][-\+]\d{2,3}' # pattern for Pilot size/waist
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
            pat = (r'(-?\d\.\d{4,6}[Ee][-\+]\d{2,3}\s*)' + r'{{{num}}}'
                   .format(num=grid_x))
            start_line = _getFirstLineOfInterest(line_list, pat)
            powerGrid = _get2DList(line_list, start_line, grid_y)

        if not keepFile:
            _deleteFile(textFileName)

        if data_is_irr: # Irradiance data
            popi = _co.namedtuple('POPinfo', ['surf', 'peakIrr', 'totPow',
                                              'fibEffSys', 'fibEffRec', 'coupling',
                                              'pilotSize', 'pilotWaist', 'pos',
                                              'rayleigh', 'gridX', 'gridY',
                                              'widthX', 'widthY' ])
            popInfo = popi(surf, peakIrr, totPow, fibEffSys, fibEffRec, coupling,
                           pilotSize, pilotWaist, pos, rayleigh,
                           grid_x, grid_y, width_x, width_y)
        else: # Phase data
            popi = _co.namedtuple('POPinfo', ['surf', 'cenPhase', 'blank',
                                              'fibEffSys', 'fibEffRec', 'coupling',
                                              'pilotSize', 'pilotWaist', 'pos',
                                              'rayleigh', 'gridX', 'gridY',
                                              'widthX', 'widthY' ])
            popInfo = popi(surf, centerPhase, None, fibEffSys, fibEffRec, coupling,
                           pilotSize, pilotWaist, pos, rayleigh,
                           grid_x, grid_y, width_x, width_y)
        if displayData:
            return (popInfo, powerGrid)
        else:
            return popInfo

    def zModifyPOPSettings(self, settingsFile, startSurf=None,
                           endSurf=None, field=None, wave=None, auto=None,
                           beamType=None, paramN=((),()), pIrr=None, tPow=None,
                           sampx=None, sampy=None, srcFile=None, widex=None,
                           widey=None, fibComp=None, fibFile=None, fibType=None,
                           fparamN=((),()), ignPol=None, pos=None, tiltx=None,
                           tilty=None):
        """Modify an existing POP settings (configuration) file

        Only those parameters that are non-None or non-zero-length (in
        case of tuples) will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        startSurf : integer, optional
            the starting surface (in General Tab)
        endSurf : integer, optional
            the end surface (in General Tab)
        field : integer, optional
            the field number (in General Tab)
        wave : integer, optional
            the wavelength number (in General Tab)
        auto : integer, optional
            simulates the pressing of the "auto" button which chooses
            appropriate X and Y widths based upon the sampling and
            other settings (in Beam Definition Tab)
        beamType : integer (0...6), optional
            0 = Gaussian Waist; 1 = Gaussian Angle; 2 = Gaussian Size +
            Angle; 3 = Top Hat; 4 = File; 5 = DLL; 6 = Multimode.
            (in Beam Definition Tab)
        paramN : 2-tuple, optional
            sets beam parameter n, for example ((1, 4),(0.1, 0.5)) sets
            parameters 1 and 4 to 0.1 and 0.5 respectively. These
            parameter names and values change depending upon the beam type
            setting. For example, for the Gaussian Waist beam, n=1 for
            Waist X, 2 for Waist Y, 3 for Decenter X, 4 for Decenter Y,
            5 for Aperture X, 6 for Aperture Y, 7 for Order X, and 8 for
            Order Y (in Beam Definition Tab)
        pIrr : float, optional
            sets the normalization by peak irradiance. It is the initial
            beam peak irradiance in power per area. It is an alternative
            to Total Power (tPow) [in Beam Definition Tab]
        tPow : float, optional
            sets the normalization by total beam power. It is the initial
            beam total power. This is an alternative to Peak Irradiance
            (pIrr) [in Beam Definition Tab]
        sampx : integer (1...10), optional
            the X direction sampling. 1 for 32; 2 for 64; 3 for 128;
            4 for 256; 5 for 512; 6 for 1024; 7 for 2048; 8 for 4096;
            9 for 8192; 10 for 16384; (in Beam Definition Tab)
        sampy : integer (1...10), optional
            the Y direction sampling. 1 for 32; 2 for 64; 3 for 128;
            4 for 256; 5 for 512; 6 for 1024; 7 for 2048; 8 for 4096;
            9 for 8192; 10 for 16384; (in Beam Definition Tab)
        srcFile : string, optional
            The file name if the starting beam is defined by a ZBF file,
            DLL, or multimode file; (in Beam Definition Tab)
        widex : float, optional
            the initial X direction width in lens units; 
            (X-Width in Beam Definition Tab)
        widey : float, optional
            the initial Y direction width in lens units;
            (Y-Width in Beam Definition Tab)
        fibComp : integer (1/0), optional
            use 1 to check the fiber coupling integral ON, 0 for OFF
            (in Fiber Data Tab)
        fibFile : string, optional
            the file name if the fiber mode is defined by a ZBF or DLL
            (in Fiber Data Tab)
        fibType : string, optional
            use the same values as ``beamType`` above, except for
            multimode which is not yet supported
            (in Fiber Data Tab)
        fparamN : 2-tuple, optional
            sets fiber parameter n, for example ((2,3),(0.5, 0.6)) sets
            parameters 2 and 3 to 0.5 and 0.6 respectively. See the hint
            for ``paramN`` (in Fiber Data Tab)
        ignPol : integer (0/1), optional
            use 1 to ignore polarization, 0 to consider polarization
            (in Fiber Data Tab)
        pos : integer (0/1), optional 
            fiber position setting. use 0 for chief ray, 1 for surface vertex
            (in Fiber Data Tab)
        tiltx : float, optional
            tilt about X in degrees (in Fiber Data Tab)
        tilty : float, optional
            tilt about Y in degrees (in Fiber Data Tab)

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetPOPSettings(), zGetPOP()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if startSurf is not None:
            sTuple.append(self.zModifySettings(dst, "POP_START", startSurf))
        if endSurf is not None:
            sTuple.append(self.zModifySettings(dst, "POP_END", endSurf))
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

    def zSetPOPSettings(self, data=0, settingsFile=None, startSurf=None,
                        endSurf=None, field=None, wave=None, auto=None,
                        beamType=None, paramN=((),()), pIrr=None, tPow=None,
                        sampx=None, sampy=None, srcFile=None, widex=None,
                        widey=None, fibComp=None, fibFile=None, fibType=None,
                        fparamN=((),()), ignPol=None, pos=None, tiltx=None,
                        tilty=None):
        """Create and set a new settings file starting from the "reset"
        settings state of the most basic lens in Zemax.

        To modify an existing POP settings file, use
        ``zModifyPOPSettings()``. Only those parameters that are non-None
        or non-zero-length (in case of tuples) will be set.

        Parameters
        ----------
        data : integer
            0 = irradiance, 1 = phase
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of settings
            file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string "_pyzdde_POP.CFG" will be created in
            the same directory as the lens file and returned
        startSurf : integer, optional
            the starting surface (in General Tab)
        endSurf : integer, optional
            the end surface (in General Tab)
        field : integer, optional
            the field number (in General Tab)
        wave : integer, optional
            the wavelength number (in General Tab)
        auto : integer, optional
            simulates the pressing of the "auto" button which chooses
            appropriate X and Y widths based upon the sampling and
            other settings (in Beam Definition Tab)
        beamType : integer (0...6), optional
            0 = Gaussian Waist; 1 = Gaussian Angle; 2 = Gaussian Size +
            Angle; 3 = Top Hat; 4 = File; 5 = DLL; 6 = Multimode.
            (in Beam Definition Tab)
        paramN : 2-tuple, optional
            sets beam parameter n, for example ((1, 4),(0.1, 0.5)) sets
            parameters 1 and 4 to 0.1 and 0.5 respectively. These
            parameter names and values change depending upon the beam type
            setting. For example, for the Gaussian Waist beam, n=1 for
            Waist X, 2 for Waist Y, 3 for Decenter X, 4 for Decenter Y,
            5 for Aperture X, 6 for Aperture Y, 7 for Order X, and 8 for
            Order Y (in Beam Definition Tab)
        pIrr : float, optional
            sets the normalization by peak irradiance. It is the initial
            beam peak irradiance in power per area. It is an alternative
            to Total Power (tPow) [in Beam Definition Tab]
        tPow : float, optional
            sets the normalization by total beam power. It is the initial
            beam total power. This is an alternative to Peak Irradiance
            (pIrr) [in Beam Definition Tab]
        sampx : integer (1...10), optional
            the X direction sampling. 1 for 32; 2 for 64; 3 for 128;
            4 for 256; 5 for 512; 6 for 1024; 7 for 2048; 8 for 4096;
            9 for 8192; 10 for 16384; (in Beam Definition Tab)
        sampy : integer (1...10), optional
            the Y direction sampling. 1 for 32; 2 for 64; 3 for 128;
            4 for 256; 5 for 512; 6 for 1024; 7 for 2048; 8 for 4096;
            9 for 8192; 10 for 16384; (in Beam Definition Tab)
        srcFile : string, optional
            The file name if the starting beam is defined by a ZBF file,
            DLL, or multimode file; (in Beam Definition Tab)
        widex : float, optional
            the initial X direction width in lens units; 
            (X-Width in Beam Definition Tab)
        widey : float, optional
            the initial Y direction width in lens units;
            (Y-Width in Beam Definition Tab)
        fibComp : integer (1/0), optional
            use 1 to check the fiber coupling integral ON, 0 for OFF
            (in Fiber Data Tab)
        fibFile : string, optional
            the file name if the fiber mode is defined by a ZBF or DLL
            (in Fiber Data Tab)
        fibType : string, optional
            use the same values as ``beamType`` above, except for
            multimode which is not yet supported
            (in Fiber Data Tab)
        fparamN : 2-tuple, optional
            sets fiber parameter n, for example ((2,3),(0.5, 0.6)) sets
            parameters 2 and 3 to 0.5 and 0.6 respectively. See the hint
            for ``paramN`` (in Fiber Data Tab)
        ignPol : integer (0/1), optional
            use 1 to ignore polarization, 0 to consider polarization
            (in Fiber Data Tab)
        pos : integer (0/1), optional 
            fiber position setting. use 0 for chief ray, 1 for surface vertex
            (in Fiber Data Tab)
        tiltx : float, optional
            tilt about X in degrees (in Fiber Data Tab)
        tilty : float, optional
            tilt about Y in degrees (in Fiber Data Tab)

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyPOPSettings()`` functions
        2. The function creates settings file ending with '_pyzdde_POP.CFG'
           in order to prevent overwritting any existing settings file not
           created by pyzdde for POP.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetPOP(), zModifyPOPSettings()
        """
        # Create a settings file with "reset" settings
        global _pDir
        if data == 1:
            clean_cfg = 'RESET_SETTINGS_POP_PHASE.CFG'
        else:
            clean_cfg = 'RESET_SETTINGS_POP_IRR.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)

        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_POP.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyPOPSettings(dst, startSurf, endSurf, field, wave, auto,
                                    beamType, paramN, pIrr, tPow, sampx, sampy,
                                    srcFile, widex, widey, fibComp, fibFile,
                                    fibType, fparamN, ignPol, pos, tiltx, tilty)
            return dst

    # FFT and Huygens PSF, MTF analysis functions
    def zGetPSFCrossSec(self, which='fft', settingsFile=None, txtFile=None,
                        keepFile=False, timeout=120):
        """Returns the cross-section data of FFT or Huygens PSF analysis

        Parameters
        ----------
        which : string, optional
            if 'fft' (default), then the FFT PSF cross-section data is
            returned;
            if 'huygens', then the Huygens PSF cross-section data is
            returned;
        settingsFile : string, optional
            * if passed, the FFT/Huygens PSF analysis will be called with
              the given configuration file (settings);
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens file post-fixed with
              "_pyzdde_FFTPSFCS.CFG"/"_pyzdde_HUYGENSPSFCS.CFG" is present,
              the settings from this file will be used;
            * if no ``settingsFile`` and no file name post-fixed with
              "_pyzdde_FFTPSFCS.CFG"/"_pyzdde_HUYGENSPSFCS.CFG" is found,
              but a config file with the same name as the lens file is
              present, the settings from that file will be used;
            * if no settings file is found, then a default settings will
              be used
        txtFile : string, optional
            if passed, the PSF analysis text file will be named such.
            Pass a specific txtFile if you want to dump the file into
            a separate directory.
        keepFile : bool, optional
            if ``False`` (default), the PSF text file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the PSF text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds. Note that Huygens PSF calculations
            may take few minutes to complete

        Returns
        -------
        indices : list
            row index of the data
        position : list
            position in microns
        value : list
            the value of the FFT/Huygens based PSF

        Notes
        -----
        The function doesn't check for inconsistencies of results. In
        most cases, if not all cases, the ``indices``, ``position``, and
        ``value`` lists should be of the same length.

        See Also
        --------
        zModifyFFTPSFCrossSecSettings(), zSetFFTPSFCrossSecSettings(),
        zModifyHuygensPSFCrossSecSettings(), zSetHuygensPSFCrossSecSettings()
        """
        if which=='huygens':
            anaType = 'Hcs'
        else:
            anaType = 'Pcs'
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, anaType)
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, anaType, cfgFile, getTextFlag,
                                timeout)
        assert ret == 0
        line_list = _readLinesFromFile(_openFile(textFileName))
        # Get Image grid size
        img_grid_line = line_list[_getFirstLineOfInterest(line_list,
                                 'Image grid size')]
        _, img_grid_y = [int(i) for i in _re.findall(r'\d{2,5}', img_grid_line)]
        pat = (r'\d{1,5}\s*(-?\d{1,3}\.\d{4,6}\s*)' + r'{{{num}}}'.format(num=2))
        start_line = _getFirstLineOfInterest(line_list, pat)
        data_mat = _get2DList(line_list, start_line, img_grid_y*2 + 1)
        data_matT = _transpose2Dlist(data_mat)
        indices = [int(i) for i in data_matT[0]]
        position = data_matT[1]
        value = data_matT[2]
        if not keepFile:
            _deleteFile(textFileName)
        return indices, position, value

    def zGetPSF(self, which='fft', settingsFile=None, txtFile=None,
                keepFile=False, timeout=120):
        """Returns FFT or Huygens PSF data

        Parameters
        ----------
        which : string, optional
            if 'fft' (default), then the FFT PSF data is returned;
            if 'huygens', then the Huygens PSF data is returned;
        settingsFile : string, optional
            * if passed, the FFT/Huygens PSF analysis will be called with
              the given configuration file (settings);
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens-file post-fixed with
              "_pyzdde_FFTPSF.CFG"/"_pyzdde_HUYGENSPSF.CFG"is present, the
              settings from this file will be used;
            * if no ``settingsFile`` and no file-name post-fixed with
              "_pyzdde_FFTPSF.CFG"/"_pyzdde_HUYGENSPSF.CFG" is found, but
              a config file with the same name as the lens file is present,
              the settings from that file will be used;
            * if no settings file is found, then a default settings will
              be used
        txtFile : string, optional
            if passed, the PSF analysis text file will be named such.
            Pass a specific txtFile if you want to dump the file into
            a separate directory.
        keepFile : bool, optional
            if ``False`` (default), the PSF text file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the PSF text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds. Note that Huygens PSF/MTF calculations with
            ``pupil_sample`` and/or ``image_sample`` greater than 4
            usually take several minutes to complete

        Returns
        -------
        psfInfo : named tuple
            meta data about the PSF analysis data, such as data spacing
            (microns), data area (microns wide), pupil and image grid
            sizes, center point, and center/reference coordinate information
        psfGridData : 2D list
            the two-dimensional list of the PSF data

        See Also
        --------
        zModifyFFTPSFSettings(), zSetFFTPSFSettings(),
        zModifyHuygensPSFSettings(), zSetHuygensPSFSettings()
        """
        if which=='huygens':
            anaType = 'Hps'
        else:
            anaType = 'Fps'
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, anaType)
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, anaType, cfgFile, getTextFlag,
                                timeout)
        assert ret == 0
        line_list = _readLinesFromFile(_openFile(textFileName))

        # Meta data
        data_spacing_line = line_list[_getFirstLineOfInterest(line_list, 'Data spacing')]
        data_spacing = float(_re.search(r'\d{1,3}\.\d{2,6}', data_spacing_line).group())
        data_area_line = line_list[_getFirstLineOfInterest(line_list, 'Data area')]
        data_area = float(_re.search(r'\d{1,5}\.\d{2,6}', data_area_line).group())
        if which=='huygens':
            ctr_ref_line = line_list[_getFirstLineOfInterest(line_list, 'Center coordinates')]
        else:
            ctr_ref_line = line_list[_getFirstLineOfInterest(line_list, 'Reference Coordinates')]
        ctr_ref_x, ctr_ref_y = [float(i) for i in _re.findall('-?\d\.\d{4,10}[Ee][-\+]\d{2,3}', ctr_ref_line)]
        img_grid_line = line_list[_getFirstLineOfInterest(line_list, 'Image grid size')]
        img_grid_x, img_grid_y = [int(i) for i in _re.findall(r'\d{2,5}', img_grid_line)]
        pupil_grid_line = line_list[_getFirstLineOfInterest(line_list, 'Pupil grid size')]
        pupil_grid_x, pupil_grid_y = [int(i) for i in _re.findall(r'\d{2,5}', pupil_grid_line)]
        center_point_line = line_list[_getFirstLineOfInterest(line_list, 'Center point')]
        center_point_x, center_point_y = [int(i) for i in _re.findall(r'\d{2,5}', center_point_line)]

        # The 2D data
        pat = (r'(-?\d\.\d{4,6}[Ee][-\+]\d{2,3}\s*)' + r'{{{num}}}'
               .format(num=img_grid_x))
        start_line = _getFirstLineOfInterest(line_list, pat)
        psfGridData = _get2DList(line_list, start_line, img_grid_y)

        if which=='huygens':
                psfi = _co.namedtuple('PSFinfo', ['dataSpacing', 'dataArea', 'pupilGridX',
                                                  'pupilGridY', 'imgGridX', 'imgGridY',
                                                  'centerPtX', 'centerPtY',
                                                  'centerCoordX', 'centerCoordY'])
        else:
                psfi = _co.namedtuple('PSFinfo', ['dataSpacing', 'dataArea', 'pupilGridX',
                                                  'pupilGridY', 'imgGridX', 'imgGridY',
                                                  'centerPtX', 'centerPtY',
                                                  'refCoordX', 'refCoordY'])

        psfInfo = psfi(data_spacing, data_area, pupil_grid_x, pupil_grid_y,
                               img_grid_x, img_grid_y, center_point_x, center_point_y,
                               ctr_ref_x, ctr_ref_y)

        if not keepFile:
            _deleteFile(textFileName)
        return (psfInfo, psfGridData)

    def zModifyFFTPSFCrossSecSettings(self, settingsFile, dtype=None, row=None,
                                      sample=None, wave=None, field=None,
                                      pol=None, norm=None, scale=None):
        """Modify an existing FFT PSF Cross section analysis settings
        (configuration) file

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        dtype : integer (0-9), optional
            0 = x-linear, 1 = y-linear, 2 = x-log, 3 = y-log, 4 = x-phase,
            5 = y-phase, 6 = x-real, 7 = y-real, 8 = x-imaginary,
            9 = y-imaginary.
        row : integer, optional
            the row number (for x scan) or column number (for y scan) or
            use 0 for center.
        sample : integer, optional
            the sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128; 4 = 256x256;
            5 = 512x512; 6 = 1024x1024; 7 = 2048x2048; 8 = 4096x4096;
            9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        norm : integer (0/1), optional
            normalization. 0 for unnormalized, 1 for unity normalization
        scale : float, optional
            the plot scale

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetFFTPSFCrossSecSettings() :
            to create and set FFT PSF Crosssection settings
        zGetPSFCrossSec(),
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_TYPE", dtype))
        if row is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_ROW", row))
        if sample is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_SAMP", sample))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_FIELD", field))
        if pol is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_POLARIZATION", pol))
        if norm is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_NORMALIZE", norm))
        if scale is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_PLOTSCALE", scale))
        return tuple(sTuple)

    def zSetFFTPSFCrossSecSettings(self, settingsFile=None, dtype=None, row=None,
                                   sample=None, wave=None, field=None, pol=None,
                                   norm=None, scale=None):
        """create and set a new FFT PSF Crosssection settings file starting
        from the "reset" settings state of the most basic lens in Zemax

        To modify an existing FFT PSF Crosssection settings file, use
        ``zModifyFFTPSFCrossSecSettings()``. Only those parameters with
        non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_FFTPSFCS.CFG' will be created
            in the same directory as the lens file and returned
        dtype : integer (0-9), optional
            0 = x-linear, 1 = y-linear, 2 = x-log, 3 = y-log, 4 = x-phase,
            5 = y-phase, 6 = x-real, 7 = y-real, 8 = x-imaginary,
            9 = y-imaginary.
        row : integer, optional
            the row number (for x scan) or column number (for y scan) or
            use 0 for center.
        sample : integer, optional
            the sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128; 4 = 256x256;
            5 = 512x512; 6 = 1024x1024; 7 = 2048x2048; 8 = 4096x4096;
            9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        norm : integer (0/1), optional
            normalization. 0 for unnormalized, 1 for unity normalization
        scale : float, optional
            the plot scale

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyFFTPSFCrossSecSettings()``
           functions
        2. The function creates settings file ending with
           '_pyzdde_FFTPSFCS.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for FFT PSF Cross
           section analysis.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetPSFCrossSec(), zModifyFFTPSFCrossSecSettings()
        """
        clean_cfg = 'RESET_SETTINGS_FFTPSFCS.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_FFTPSFCS.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyFFTPSFCrossSecSettings(dst, dtype, row, sample, wave,
                                               field, pol, norm, scale)
            return dst

    def zModifyFFTPSFSettings(self, settingsFile, dtype=None, sample=None,
                              wave=None, field=None, surf=None, pol=None,
                              norm=None, imgDelta=None):
        """Modify an existing FFT PSF analysis settings (configuration)
        file

        Only those parameters that are non-None will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        dtype : integer (0-4), optional
            0 = linear, 1 = log, 2 = phase, 3 = real, 4 = imaginary.
        sample : integer, optional
            the (pupil) sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number
        surf : integer, optional
            the surface number. Use 0 for image
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        norm : integer (0/1), optional
            normalization. 0 for unnormalized, 1 for unity normalization
        imgDelta : float, optional
            the image point spacing in micrometers

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        Notes
        -----
        See the notes of ``zSetFFTPSFSettings()``

        See Also
        --------
        zSetFFTPSFSettings(), zGetPSF()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_TYPE", dtype))
        if sample is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_SAMP", sample))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_FIELD", field))
        if surf is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_SURFACE", surf))
        if pol is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_POLARIZATION", pol))
        if norm is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_NORMALIZE", norm))
        if imgDelta is not None:
            sTuple.append(self.zModifySettings(dst, "PSF_IMAGEDELTA", imgDelta))
        return tuple(sTuple)

    def zSetFFTPSFSettings(self, settingsFile=None, dtype=None, sample=None,
                           wave=None, field=None, surf=None, pol=None,
                           norm=None, imgDelta=None):
        """create and set a new FFT PSF analysis settings file starting
        from the "reset" settings state of the most basic lens in Zemax

        To modify an existing FFT PSF settings file, use
        ``zModifyFFTPSFSettings()``. Only those parameters that are
        non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_FFTPSF.CFG' will be created
            in the same directory as the lens file and returned
        dtype : integer (0-4), optional
            0 = linear, 1 = log, 2 = phase, 3 = real, 4 = imaginary.
        sample : integer, optional
            the (pupil) sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number
        surf : integer, optional
            the surface number. Use 0 for image
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        norm : integer (0/1), optional
            normalization. 0 for unnormalized, 1 for unity normalization
        imgDelta : float, optional
            the image point spacing in micrometers

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Currently, Zemax doesn't provide a way to change the image
           sampling parameter for this function. It seems that the image
           sampling value is set to twice the value set for pupil sampling.
        2. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyFFTPSFSettings()`` functions
        3. The function creates settings file ending with
           '_pyzdde_FFTPSF.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for FFT PSF.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetPSF(), zModifyFFTPSFSettings()
        """
        clean_cfg = 'RESET_SETTINGS_FFTPSF.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_FFTPSF.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyFFTPSFSettings(dst, dtype, sample, wave, field, surf, pol,
                                       norm, imgDelta)
            return dst

    def zModifyHuygensPSFCrossSecSettings(self, settingsFile, pupilSample=None,
                                          imgSample=None, wave=None, field=None,
                                          imgDelta=None, dtype=None):
        """Modify an existing Huygens PSF Cross section analysis settings
        (configuration) file

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic
        field : integer, optional
            the field number
        imgDelta : float, optional
            the image point spacing in micrometers
        dtype : integer (0-9), optional
            0 = x-linear, 1 = y-log, 2 = y-linear, 3 = y-log, 4 = x-real,
            5 = y-real, 6 = x-imaginary, 7 = y-imaginary, 8 = x-phase,
            9 = y-phase.

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetHuygensPSFCrossSecSettings() :
            to create and set Huygens PSF Crosssection settings
        zGetPSFCrossSec(),
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if pupilSample is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_PUPILSAMP", pupilSample))
        if imgSample is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_IMAGESAMP", imgSample))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_FIELD", field))
        if imgDelta is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_IMAGEDELTA", imgDelta))
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "HPC_TYPE", dtype))
        return tuple(sTuple)

    def zSetHuygensPSFCrossSecSettings(self, settingsFile=None, pupilSample=None,
                                       imgSample=None, wave=None, field=None,
                                       imgDelta=None, dtype=None):
        """create and set a new Huygens PSF Crosssection settings file
        starting from the "reset" settings state of the most basic lens in
        Zemax.

        To modify an existing Huygens PSF Crosssection settings file, use
        ``zModifyHuygensPSFCrossSecSettings()``. Only those parameters
        with non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_HUYGENSPSFCS.CFG' will be
            created in the same directory as the lens file and returned
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic
        field : integer, optional
            the field number
        imgDelta : float, optional
            the image point spacing in micrometers
        dtype : integer (0-9), optional
            0 = x-linear, 1 = y-log, 2 = y-linear, 3 = y-log, 4 = x-real,
            5 = y-real, 6 = x-imaginary, 7 = y-imaginary, 8 = x-phase,
            9 = y-phase.

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyHuygensPSFCrosSecSettings()``
           functions
        2. The function creates settings file ending with
           '_pyzdde_HUYGENSPSFCS.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for Huygens PSF
           Crosssection analysis.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetPSFCrossSec(), zModifyHuygensPSFCrossSecSettings()
        """
        clean_cfg = 'RESET_SETTINGS_HUYGENSPSFCS.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_HUYGENSPSFCS.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyHuygensPSFCrossSecSettings(dst, pupilSample, imgSample,
                                                   wave, field, imgDelta, dtype)
            return dst

    def zModifyHuygensPSFSettings(self, settingsFile, pupilSample=None,
                                  imgSample=None, wave=None, field=None,
                                  imgDelta=None, dtype=None):
        """Modify an existing Huygens PSF analysis settings (configuration)
        file

        Only those parameters that are non-None will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic
        field : integer, optional
            the field number
        imgDelta : float, optional
            the image point spacing in micrometers
        dtype : integer (0-8), optional
            0 = linear, 1 = log -1, 2 = log -2, 3 = log -3, 4 = log -4,
            5 = log -5, 6 = real, 7 = imaginary, 8 = phase.

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetHuygensPSFSettings(), zGetPSF()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if pupilSample is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_PUPILSAMP", pupilSample))
        if imgSample is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_IMAGESAMP", imgSample))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_FIELD", field))
        if imgDelta is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_IMAGEDELTA", imgDelta))
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "HPS_TYPE", dtype))
        return tuple(sTuple)

    def zSetHuygensPSFSettings(self, settingsFile=None, pupilSample=None,
                               imgSample=None, wave=None, field=None,
                               imgDelta=None, dtype=None):
        """create and set a new Huygens PSF analysis settings file starting
        from the "reset" settings state of the most basic lens in Zemax

        To modify an existing Huygens PSF settings file, use
        ``zModifyHuygensPSFSettings()``. Only those parameters that are
        non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_HUYGENSPSF.CFG' will be
            created in the same directory as the lens file and returned
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic
        field : integer, optional
            the field number
        imgDelta : float, optional
            the image point spacing in micrometers
        dtype : integer (0-8), optional
            0 = linear, 1 = log -1, 2 = log -2, 3 = log -3, 4 = log -4,
            5 = log -5, 6 = real, 7 = imaginary, 8 = phase.

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyHuygensPSFSettings()``
           functions
        2. The function creates settings file ending with
           '_pyzdde_HUYGENSPSF.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for Huygens PSF
           analysis.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetPSF(), zModifyHuygensPSFSettings()
        """
        clean_cfg = 'RESET_SETTINGS_HUYGENSPSF.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_HUYGENSPSF.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyHuygensPSFSettings(dst, pupilSample, imgSample, wave,
                                           field, imgDelta, dtype)
            return dst

    def zGetMTF(self, which='fft', settingsFile=None, txtFile=None,
                keepFile=False, timeout=120):
        """Returns FFT or Huygens MTF data

        Parameters
        ----------
        which : string, optional
            if 'fft' (default), then the FFT MTF data is returned;
            if 'huygens', then the Huygens MTF data is returned;
        settingsFile : string, optional
            * if passed, the FFT/Huygens MTF analysis will be called with
              the given configuration file (settings);
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens-file post-fixed with
              "_pyzdde_FFTMTF.CFG"/"_pyzdde_HUYGENSMTF.CFG"is present, the
              settings from this file will be used;
            * if no ``settingsFile`` and no file name post-fixed with
              "_pyzdde_FFTMTF.CFG"/"_pyzdde_HUYGENSMTF.CFG" is found, but
              a config file with the same name as the lens file is present,
              the settings from that file will be used;
            * if no settings file is found, then a default settings will
              be used
        txtFile : string, optional
            if passed, the MTF analysis text file will be named such.
            Pass a specific txtFile if you want to dump the file into
            a separate directory.
        keepFile : bool, optional
            if ``False`` (default), the MTF text file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the MTF text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds. Note that Huygens PSF/MTF calculations with
            ``pupil_sample`` and/or ``image_sample`` greater than 4
            usuallly take several minutes to complete

        Returns
        -------
        mtfs : tuple of tuples
            the tuple contains MTF data for the number of fields defined
            in the MTF analysis configuration/settings. The len of the
            tuple equals the number of fields. Each sub-tuple is a named
            tuple that contains Spatial frequency, Tangential, and
            Sagittal MTF values. The information can be retrieved as shown
            in the example below.

        Examples
        --------
        The following example plots the MTFs for each defined field points
        of a Zemax lens

        >>> mtfs = ln.zGetMTF()
        >>> for field, mtf in enumerate(mtfs):
        >>>     plt.plot(mtf.SpatialFreq, mtf.Tangential, label='F-{}, T'.format(field + 1))
        >>>     plt.plot(mtf.SpatialFreq, mtf.Sagittal, label='F-{}, S'.format(field + 1))
        >>>     plt.xlabel('Spatial Frequency in cycles per mm')
        >>>     plt.ylabel('Modulus of the OTF')
        >>>     plt.grid('on')
        >>> plt.legend(frameon=False)
        >>> plt.show()

        See Also
        --------
        zModifyFFTMTFSettings(), zSetFFTMTFSettings(),
        zModifyHuygensMTFSettings(), zSetHuygensMTFSettings()
        """
        if which=='huygens':
            anaType = 'Hmf'
        else:
            anaType = 'Mtf'
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, anaType)
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, anaType, cfgFile, getTextFlag,
                                timeout)
        assert ret == 0
        line_list = _readLinesFromFile(_openFile(textFileName))
        pat = r'Field:\s-?\d{1,3}\.\d{1,5},?\s?'
        fields = _getRePatPosInLineList(line_list, pat)
        if len(fields) > 1:
            data_start_pos = [p + 2 for p in fields]
            data_len = [fields[1] - fields[0] - 3]*len(fields)
        else:
            data_start_pos = [fields[0] + 2,]
            data_len = [len(line_list) - data_start_pos[0],]
        mtfs = []
        mtf = _co.namedtuple('MTF', ['SpatialFreq', 'Tangential', 'Sagittal'])
        for start, length in zip(data_start_pos, data_len):
            data_mat = _get2DList(line_list, start, length)
            data_matT = _transpose2Dlist(data_mat)
            spat_freq = data_matT[0]
            mtf_tang = data_matT[1]
            mtf_sagi = data_matT[2]
            mtfs.append(mtf(spat_freq, mtf_tang, mtf_sagi))
        if not keepFile:
            _deleteFile(textFileName)
        return tuple(mtfs)

    def zModifyFFTMTFSettings(self, settingsFile, sample=None, wave=None,
                              field=None, dtype=None, surf=None, maxFreq=None,
                              showDiff=None, pol=None, useDash=None):
        """Modify an existing FFT MTF analysis settings (configuration)
        file

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        sample : integer, optional
            the sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128; 4 = 256x256;
            5 = 512x512; 6 = 1024x1024; 7 = 2048x2048; 8 = 4096x4096;
            9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number, 0 for all
        dtype : integer (0-4), optional
            0 = modulation, 1 = real, 2 = imaginary, 3 = phase, 4 = square
            wave.
        surf : integer, optional
            the surface number. Use 0 for image
        maxFreq : real, optional
            the maximum frequency, use 0 for default
        showDiff : integer (0/1)
            show diffraction limit, 0 for no, 1 for yes
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        useDash : integer (0/1)
            use dashes, 0 for no, 1 for yes

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetFFTMTFSettings() :
            to create and set FFT MTF settings/configuration file
        zGetMTF(),
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if sample is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_SAMP", sample))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_FIELD", field))
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_TYPE", dtype))
        if surf is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_SURF", surf))
        if maxFreq is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_MAXF", maxFreq))
        if showDiff is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_SDLI", showDiff))
        if pol is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_POLAR", pol))
        if useDash is not None:
            sTuple.append(self.zModifySettings(dst, "MTF_DASH", useDash))
        return tuple(sTuple)

    def zSetFFTMTFSettings(self, settingsFile=None, sample=None, wave=None,
                           field=None, dtype=None, surf=None, maxFreq=None,
                           showDiff=None, pol=None, useDash=None):
        """create and set a new FFT MTF analysis settings file starting
        from the "reset" settings state of the most basic lens in Zemax

        To modify an existing FFT MTF settings file, use
        ``zModifyFFTMTFSettings()``. Only those parameters that are
        non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_FFTMTF.CFG' will be created
            in the same directory as the lens file and returned
        sample : integer, optional
            the sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128; 4 = 256x256;
            5 = 512x512; 6 = 1024x1024; 7 = 2048x2048; 8 = 4096x4096;
            9 = 8192x8192; 10 = 16384x16384;
        wave : integer, optional
            the wavelength number, use 0 for polychromatic.
        field : integer, optional
            the field number, 0 for all
        dtype : integer (0-4), optional
            0 = modulation, 1 = real, 2 = imaginary, 3 = phase, 4 = square
            wave.
        surf : integer, optional
            the surface number. Use 0 for image
        maxFreq : real, optional
            the maximum frequency, use 0 for default
        showDiff : integer (0/1)
            show diffraction limit, 0 for no, 1 for yes
        pol : integer (0/1), optional
            the polarization. 0 for unpolarized, 1 for polarized.
        useDash : integer (0/1)
            use dashes, 0 for no, 1 for yes

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyFFTMTFSettings()`` functions
        2. The function creates settings file ending with
           '_pyzdde_FFTMTF.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for FFT MTF.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetMTF(), zModifyFFTMTFSettings()
        """
        clean_cfg = 'RESET_SETTINGS_FFTMTF.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_FFTMTF.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyFFTMTFSettings(dst, sample, wave, field, dtype, surf,
                                       maxFreq, showDiff, pol, useDash)
            return dst

    def zModifyHuygensMTFSettings(self, settingsFile, pupilSample=None,
                                  imgSample=None, imgDelta=None, config=None,
                                  wave=None, field=None, dtype=None, maxFreq=None,
                                  pol=None, useDash=None):
        """Modify an existing Huygens MTF analysis settings (configuration)
        file

        Only those parameters that are non-None will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgDelta : float, optional
            the image point spacing in micrometers
        config : integer, optional
            the configuration number. Use 0 for all, 1 for current, etc.
        wave : integer, optional
            the wavelength number. Use 0 for polychromatic
        field : integer, optional
            the field number
        dtype : integer, optional
            the data type. Currently only 0 is supported
        maxFreq : float, optional
            the maximum spatial frequency
        pol : integer, optional
            polarization. 0 for no, 1 for yes
        useDash : integer, optional
            use dashes. 0 for no, 1 for yes

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetHuygensMTFSettings(), zGetMTF()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if pupilSample is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_PUPILSAMP", pupilSample))
        if imgSample is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_IMAGESAMP", imgSample))
        if imgDelta is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_IMAGEDELTA", imgDelta))
        if config is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_CONFIG", config))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_FIELD", field))
        if dtype is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_TYPE", dtype))
        if maxFreq is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_MAXF", maxFreq))
        if pol is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_POLAR", pol))
        if useDash is not None:
            sTuple.append(self.zModifySettings(dst, "HMF_DASH", useDash))
        return tuple(sTuple)

    def zSetHuygensMTFSettings(self, settingsFile=None, pupilSample=None,
                               imgSample=None, imgDelta=None, config=None,
                               wave=None, field=None, dtype=None, maxFreq=None,
                               pol=None, useDash=None):
        """create and set a new Huygens MTF analysis settings file starting
        from the "reset" settings state of the most basic lens in Zemax

        To modify an existing Huygens MTF settings file, use
        ``zModifyHuygensMTFSettings()``. Only those parameters that are
        non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_HUYGENSMTF.CFG' will be
            created in the same directory as the lens file and returned
        pupilSample : integer, optional
            the pupil sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgSample : integer, optional
            the image sampling. 1 = 32x32; 2 = 64x64; 3 = 128x128;
            4 = 256x256; 5 = 512x512; 6 = 1024x1024; 7 = 2048x2048;
            8 = 4096x4096; 9 = 8192x8192; 10 = 16384x16384;
        imgDelta : float, optional
            the image point spacing in micrometers
        config : integer, optional
            the configuration number. Use 0 for all, 1 for current, etc.
        wave : integer, optional
            the wavelength number. Use 0 for polychromatic
        field : integer, optional
            the field number
        dtype : integer, optional
            the data type. Currently only 0 is supported
        maxFreq : float, optional
            the maximum spatial frequency
        pol : integer, optional
            polarization. 0 for no, 1 for yes
        useDash : integer, optional
            use dashes. 0 for no, 1 for yes

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyHuygensMTFSettings()``
           functions
        2. The function creates settings file ending with
           '_pyzdde_HUYGENSMTF.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for Huygens MTF
           analysis.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetMTF(), zModifyHuygensMTFSettings()
        """
        clean_cfg = 'RESET_SETTINGS_HUYGENSMTF.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_HUYGENSMTF.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyHuygensMTFSettings(dst, pupilSample, imgSample, imgDelta,
                                           config, wave, field, dtype, maxFreq,
                                           pol, useDash)
            return dst

    # Image simulation functions
    def zGetImageSimulation(self, settingsFile=None, txtFile=None, keepFile=False,
                            timeout=120):
        """Returns image simulation analysis results

        Parameters
        ----------
        settingsFile : string, optional
            * if passed, the image simulation analysis will be called with
              the given configuration file (settings);
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens-file post-fixed with
              "_pyzdde_IMGSIM.CFG" is present, the settings from this file
              will be used;
            * if no ``settingsFile`` and no file-name post-fixed with
              "_pyzdde_IMGSIM.CFG" is found, but a config file with the
              same name as the lens file is present, the settings from
              that file will be used;
            * if no settings file is found, then a default settings will
              be used
        txtFile : string, optional
            if passed, the image simulation analysis text file will be
            named such. Pass a specific txtFile if you want to dump the
            file into a separate directory.
        keepFile : bool, optional
            if ``False`` (default), the image simulation text file will be
            deleted after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds.

        Returns
        -------
        imgInfo : named tuple
            meta data about the image analysis data containing 'xpix',
            'ypix', 'objHeight', 'fieldPos', 'imgW', and 'imgH'. PSF 
            Grid data doesn't have `imgW` and `imgH` and Source bitmap 
            image data only has `xpix` and `ypix`. 
        imgData : 3D list
            the 3D list containing the RGB values of the output image.
            The first dimension of ``imgData`` represents height (rows),
            the second dimension represents width (cols), and the third
            dimension represents the channel (r, g, b)

        Examples
        --------
        In the following example the image simulation function is called with 
        default arguments, and the returned data is plotted using matplotlib's 
        imshow function after converting the data into a Numpy (np) array.

        >>> cfgfile = ln.zSetImageSimulationSettings(image='RGB_CIRCLES.BMP', height=1)
        >>> img_info, img_data = ln.zGetImageSimulationData(settingsFile=cfgfile)
        >>> img = np.array(img_data, dtype='uint8')
        >>> fig, ax = plt.subplots(1,1, figsize=(10, 8))
        >>> if len(img_info)==6: # image simulation data
        >>>     bottom, top = -img_info.imgH/2, img_info.imgH/2
        >>>     left, right = -img_info.imgW/2, img_info.imgW/2
        >>>     extent=[left, right, bottom, top]
        >>>     xl, yl = 'Image width (mm)', 'Image height (mm)'
        >>>     t = 'Simulated Image'
        >>> elif len(img_info)==4: # psf grid data
        >>>     bottom, top = -img_info.objHeight/2, img_info.objHeight/2
        >>>     aratio = img_info.xpix/img_info.ypix
        >>>     left, right =  bottom*aratio, top*aratio
        >>>     extent=[left, right, bottom, top]
        >>>     xl, yl = 'Field width (mm)', 'Field height (mm)'
        >>>     t = 'PSF Grid at field pos {:2.2f}'.format(img_info.fieldPos)
        >>> else: # source bitmap
        >>>     extent = [0, img_info.xpix, 0, img_info.ypix]
        >>>     xl = '{} pixels wide'.format(img_info.xpix) 
        >>>     yl = '{} pixels high'.format(img_info.ypix)
        >>>     t = 'Source Bitmap'
        >>> ax.imshow(img, extent=extent, interpolation='none')
        >>> ax.set_xlabel(xl); ax.set_ylabel(yl); ax.set_title(t)
        >>> plt.show()

        Notes
        ----- 
        It is recommended that a settings files is first generated using the 
        ``zSetImageSimulationSettings()`` functions prior to calling 
        ``zGetImageSimulation()``.

        See Also
        --------
        zModifyImageSimulationSettings(), zSetImageSimulationSettings()
        """
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, 'Sim')
        textFileName, cfgFile, getTextFlag = settings

        ret = self.zGetTextFile(textFileName, 'Sim', cfgFile, getTextFlag,
                                timeout)
        assert ret == 0, 'zGetTextFile() returned error code {}'.format(ret)
        line_list = _readLinesFromFile(_openFile(textFileName))

        # Meta data
        data = None
        data_line = line_list[_getFirstLineOfInterest(line_list, 'Data')]
        dataType = data_line.split(':')[1].strip()
        if dataType == 'Simulated Image':
            data = 'img'
        elif dataType == 'PSF Grid':
            data = 'psf'
        else: # source bitmap
            data = 'src'
        
        bm_ht_line = line_list[_getFirstLineOfInterest(line_list, 'Bitmap Height')]
        bm_ht = int(_re.search(r'\b\d{1,5}\b', bm_ht_line).group()) # pixels
        bm_wd_line = line_list[_getFirstLineOfInterest(line_list, 'Bitmap Width')]
        bm_wd = int(_re.search(r'\b\d{1,5}\b', bm_wd_line).group())   # pixels
        
        if data=='img' or data=='psf':
            obj_ht_line = line_list[_getFirstLineOfInterest(line_list, 'Object Height')]
            obj_ht = float(_re.search(r'\b-?\d{1,3}\.\d{1,5}\b', obj_ht_line).group())
            fld_pos_line = line_list[_getFirstLineOfInterest(line_list, 'Field position')]
            fld_pos = float(_re.search(r'\b-?\d{1,3}\.\d{1,5}\b', fld_pos_line).group())
        
        if data=='img':
            img_siz_line = line_list[_getFirstLineOfInterest(line_list, 'Image Size')]
            pat = r'\d{1,3}\.\d{4,6}'
            img_wd, img_ht = [float(i) for i in _re.findall(pat, img_siz_line)] # physical units

        if data=='img':
            img_info = _co.namedtuple('ImgSimInfo', ['xpix', 'ypix', 'objHeight',
                                      'fieldPos', 'imgW', 'imgH'])
            img_info_data = img_info._make([bm_wd, bm_ht, obj_ht, fld_pos, img_wd, img_ht])
        elif data=='psf':
            img_info = _co.namedtuple('PSFGridInfo', ['xpix', 'ypix', 'objHeight',
                                      'fieldPos'])
            img_info_data = img_info._make([bm_wd, bm_ht, obj_ht, fld_pos])
        else: # source bitmap / data = src
            img_info = _co.namedtuple('SrcImgInfo', ['xpix', 'ypix'])
            img_info_data = img_info._make([bm_wd, bm_ht])
        
        img_data = [[[0 for c in range(3)] for i in range(bm_wd)] for j in range(bm_ht)]
        r, g, b = 0, 1, 2
        pat = r'xpix\s{1,4}ypix\s{1,4}R\s{1,4}G\s{1,4}B'
        start = _getFirstLineOfInterest(line_list, pat) + 1
        for xpix in range(bm_wd):      # along width
            for ypix in range(bm_ht):  # along height
                pixel_data = line_list[start + xpix*bm_ht + ypix].split()[2:]
                pix_r, pix_g, pix_b = pixel_data
                img_data[ypix][xpix][r] = int(pix_r)
                img_data[ypix][xpix][g] = int(pix_g)
                img_data[ypix][xpix][b] = int(pix_b)
        if not keepFile:
            _deleteFile(textFileName)
        return img_info_data, img_data

    def zModifyImageSimulationSettings(self, settingsFile, image=None, height=None,
                                       over=None, guard=None, flip=None, rotate=None,
                                       wave=None, field=None, pupilSample=None,
                                       imgSample=None, psfx=None, psfy=None, aberr=None,
                                       pol=None, fixedAper=None, illum=None, showAs=None,
                                       reference=None, suppress=None, pixelSize=None,
                                       xpix=None, ypix=None, flipSimImg=None, outFile=None):
        """Modify an existing image simulation analysis settings
        (configuration) file

        Only those parameters that are non-None will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        image : string, optional
            The input file name. This should be specified without a path.
        height : float, optional
            The field height, which defines the full height of the source
            bitmap in field coordinates, may be either lens units or
            degrees, depending upon the current field definition (heights
            or angles, respectively).
        over : integer, optional, [0-6]
            Oversample value. Use 0 for None, 1 for 2X, 2 for 4x, etc.
        guard : integer, optional, [0-6]
            Guard band value. Use 0 for None, 1 for 2X, 2 for 4x, etc.
        flip : integer, optional, [0-3]
            Flip Source. Use 0 for None, 1 for top-bottom, 2 for left-right, 
            3 for top-bottom & left-right.
        rotate : integer, optional, [0-3]
            Rotate Source. Use 0 for none, 1 for 90, 2 for 180, 3 for 270.
        wave : integer, optional, 
            Wavelength. Use 0 for RGB, 1 for 1+2+3, 2 for wave #1, 3 for
            wave #2, etc.
        field : integer, optional
            Field number.
        pupilSample : integer, optional, [1-10]
            Pupil Sampling. Use 1 for 32x32, 2 for 64x64, etc.
        imgSample : integer, optional, [1-5]
            Image Sampling. Use 1 for 32x32, 2 for 64x64, etc.
        psfx, psfy : integer, optional, [1-51]
            The number of PSF grid points.
        aberr : integer, optional, [0-2]
            Use 0 for none, 1 for geometric, 2 for diffraction.
        pol : integer, optional, [0-1]
            Polarization. Use 0 for no, 1 for yes.
        fixedAper : integer, optional, [0-1]
            Apply fixed aperture? Use 0 for no, 1 for yes (apply fixed
            aperture).
        illum : integer, optional, [0-1]
            Relative illumination. Use 0 for no, 1 for yes.
        showAs : integer, optional, [0-2]
            Use 0 for Simulated Image, 1 for Source Bitmap, and 2 for PSF
            Grid.
        reference : integer, optional, [0-2]
            Use 0 for chief ray, 1 for vertex, 2 for primary chief ray.
        suppress : integer, optional, [0-1]
            Use 0 for no, 1 for yes.
        pixelSize : integer, optional
            Use 0 for default or the size in lens units.
        xpix, ypix : integer, optional
            Use 0 for default or the number of pixels.
        flipSimImg : integer, optional
            Use 0 for none, 1 for top-bottom, etc.
        outFile : string, optional
            The output file name or empty string for no output file.

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetImageSimulationSettings(), zGetImageSimulation()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if image is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_INPUTFILE", image))
        if height is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_FIELDHEIGHT", height))
        if over is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_OVERSAMPLING", over))
        if guard is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_GUARDBAND", guard))
        if flip is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_FLIP", flip))
        if rotate is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_ROTATE", rotate))
        if wave is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_WAVE", wave))
        if field is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_FIELD", field))
        if pupilSample is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_PSAMP", pupilSample))
        if imgSample is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_ISAMP", imgSample))
        if psfx is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_PSFX", psfx))
        if psfy is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_PSFY", psfy))
        if aberr is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_ABERRATIONS", aberr))
        if pol is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_POLARIZATION", pol))
        if fixedAper is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_FIXEDAPERTURES", fixedAper))
        if illum is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_USERI", illum))
        if showAs is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_SHOWAS", showAs))
        if reference is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_REFERENCE", reference))
        if suppress is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_SUPPRESS", suppress))
        if pixelSize is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_PIXELSIZE", pixelSize))
        if xpix is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_XSIZE", xpix))
        if ypix is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_YSIZE", ypix))
        if flipSimImg is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_FLIPIMAGE", flipSimImg))
        if outFile is not None:
            sTuple.append(self.zModifySettings(dst, "ISM_OUTPUTFILE", outFile))
        return tuple(sTuple)

    def zSetImageSimulationSettings(self, settingsFile=None, image=None, height=None,
                                    over=None, guard=None, flip=None, rotate=None,
                                    wave=None, field=None, pupilSample=None,
                                    imgSample=None, psfx=None, psfy=None, aberr=None,
                                    pol=None, fixedAper=None, illum=None, showAs=None,
                                    reference=None, suppress=None, pixelSize=None,
                                    xpix=None, ypix=None, flipSimImg=None, outFile=None):
        """create and set a new image simulation analysis settings file
        starting from the "reset" settings state of the most basic lens in
        Zemax

        To modify an existing image simulation analysis settings file, use
        ``zModifyImageSimulationSettings()``. Only those parameters that
        are non-None will be set

        Parameters
        ----------
        settingsFile : string, optional
            name to give to the settings file to be created. It must be
            the full file name, including path and extension of the
            settings file.
            If ``None``, then a CFG file with the name of the lens
            followed by the string '_pyzdde_IMGSIM.CFG' will be created
            in the same directory as the lens file and returned
        image : string, optional
            The input file name. This should be specified without a path.
        height : float, optional
            The field height, which defines the full height of the source
            bitmap in field coordinates, may be either lens units or
            degrees, depending upon the current field definition (heights
            or angles, respectively).
        over : integer, optional, [0-6]
            Oversample value. Use 0 for None, 1 for 2X, 2 for 4x, etc.
        guard : integer, optional, [0-6]
            Guard band value. Use 0 for None, 1 for 2X, 2 for 4x, etc.
        flip : integer, optional, [0-3]
            Flip Source. Use 0 for None, 1 for top-bottom, 2 for left-right, 
            3 for top-bottom & left-right.
        rotate : integer, optional, [0-3]
            Rotate Source. Use 0 for none, 1 for 90, 2 for 180, 3 for 270.
        wave : integer, optional, 
            Wavelength. Use 0 for RGB, 1 for 1+2+3, 2 for wave #1, 3 for
            wave #2, etc.
        field : integer, optional
            Field number.
        pupilSample : integer, optional, [1-10]
            Pupil Sampling. Use 1 for 32x32, 2 for 64x64, etc.
            i.e. [32*(2**i) for i in range(10)]
        imgSample : integer, optional, [1-5]
            Image Sampling. Use 1 for 32x32, 2 for 64x64, etc.
            i.e. [32*(2**i) for i in range(5)]
        psfx, psfy : integer, optional, [1-51]
            The number of PSF grid points.
        aberr : integer, optional, [0-2]
            Use 0 for none, 1 for geometric, 2 for diffraction.
        pol : integer, optional, [0-1]
            Polarization. Use 0 for no, 1 for yes.
        fixedAper : integer, optional, [0-1]
            Apply fixed aperture? Use 0 for no, 1 for yes (apply fixed
            aperture).
        illum : integer, optional, [0-1]
            Relative illumination. Use 0 for no, 1 for yes.
        showAs : integer, optional, [0-2]
            Use 0 for Simulated Image, 1 for Source Bitmap, and 2 for PSF
            Grid.
        reference : integer, optional, [0-2]
            Use 0 for chief ray, 1 for vertex, 2 for primary chief ray.
        suppress : integer, optional, [0-1]
            Use 0 for no, 1 for yes.
        pixelSize : integer, optional
            Use 0 for default or the size in lens units.
        xpix, ypix : integer, optional
            Use 0 for default or the number of pixels.
        flipSimImg : integer, optional
            Use 0 for none, 1 for top-bottom, etc.
        outFile : string, optional
            The output file name or empty string for no output file.

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        -----
        1. Further modifications of the settings file can be made using
           ``zModifySettings()`` or ``zModifyImageSimulationSettings()``
           functions
        2. The function creates settings file ending with
           '_pyzdde_IMGSIM.CFG' in order to prevent overwritting any
           existing settings file not created by pyzdde for image
           simulation.
           This file eventually gets deleted when ``ln.close()`` or
           ``pyz.closeLink()`` or ``ln.zDDEClose()`` is called.

        See Also
        --------
        zGetImageSimulation(), zModifyImageSimulationSettings()
        """
        clean_cfg = 'RESET_SETTINGS_IMGSIM.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_IMGSIM.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyImageSimulationSettings(dst, image, height, over, guard,
                                                flip, rotate, wave, field, pupilSample,
                                                imgSample, psfx, psfy, aberr, pol,
                                                fixedAper, illum, showAs, reference,
                                                suppress, pixelSize, xpix, ypix,
                                                flipSimImg, outFile)
            return dst

    # NSC detector viewer data
    def zGetDetectorViewer(self, settingsFile=None, displayData=False, txtFile=None,
                           keepFile=False, timeout=60):
        """Returns NSC detector viewer data. 

        Please execute `zNSCTrace()` before calling this function.   

        Parameters
        ----------
        settingsFile : string, optional
            * if passed, the detector viewer uses this configuration file;
            * if no ``settingsFile`` is passed, and config file ending
              with the same name as the lens file post-fixed with
              "_pyzdde_DVW.CFG" is present, the settings from this file
              will be used;
            * if no ``settingsFile`` and no file name post-fixed with
              "_pyzdde_DVW.CFG" is found, but a config file with the same
              name as the lens file is present, the settings from that
              file will be used;
            * if no settings file is found, then a default settings will
              be used
        displayData : bool
            if ``true`` the function returns the 1D or 2D display data as
            specified in the settings file; default is ``false``
        txtFile : string, optional
            if passed, the detector viewer data file will be named such. 
            Pass a specific ``txtFile`` if you want to dump the file into a
            separate directory.
        keepFile : bool, optional
            if ``False`` (default), the detector viewer text file will be 
            deleted after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the detector viewer text file 
            will be saved in the same directory as the lens (provided the 
            required folder access permissions are available)
        timeout : integer, optional
            timeout in seconds.   

        Return 
        ------
        dvwData : tuple 
            dvwData is a 1-tuple containing just ``dvwInfo`` (see below)
            if ``displayData`` is ``False`` (default).
            If ``displayData`` is ``True``, ``dvwData`` is a 2-tuple
            containing ``dvwInfo`` (a named tuple) and ``data``. ``data``
            is either a 2-tuple containing ``coordinate`` and ``values``
            as list elements if "Show as" is row/column cross-section, or 
            a 2D list of values otherwise.

            dvwInfo : named tuple
                surfNum : integer
                    NSCG surface number
                detNum : integer 
                    detector number 
                width, height : float
                    width and height of the detector 
                xPix, yPix : integer
                    number of pixels in x and y direction
                totHits : integer 
                    total ray hits 
                peakIrr : float or None 
                    peak irradiance (only available for Irradiance type of data) 
                totPow : float or None 
                    total power (only available for Irradiance type of data)
                smooth : integer 
                    the integer smoothing value  
                dType : string
                    the "Show Data" type 
                x, y, z, tiltX, tiltY, tiltZ : float 
                    the x, y, z positions and tilt values 
                posUnits : string 
                    position units
                units : string 
                    units  
                rowOrCol : string or None
                    indicate whether the cross-section data is a row or column 
                    cross-section 
                rowColNum : float or None
                    the row or column number for cross-section data 
                rowColVal : float or None 
                    the row or column value for cross-section data

            data : 2-tuple or 2-D list 
                if cross-section data then `data = (coordinates, values)` where, 
                `coordinates` and `values` are 1-D lists otherwise, `data` is 
                a 2-D list of grid data. Note that the coherent phase data is 
                in degrees.

        Examples
        -------- 
        >>> info = ln.zGetDetectorViewer(settingsFile)
        >>> # following line assumes row/column cross-section data 
        >>> info, coordinates, values =ln.zGetDetectorViewer(settingsFile, True)
        >>> # following line assumes 2d data 
        >>> info, gridData = zfu.zGetDetectorViewer(settingsFile, True)  

        See Also
        --------
        zSetDetectorViewerSettings(), zModifyDetectorViewerSettings()
        """
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, 'Dvw')
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, 'Dvr', cfgFile, getTextFlag,
                                timeout)
        assert ret == 0, 'zGetTextFile returned {}'.format(ret)

        pyz = _sys.modules[__name__]
        ret = _zfu.readDetectorViewerTextFile(pyz, textFileName, displayData)

        if not keepFile:
            _deleteFile(textFileName)

        return ret
        

    def zModifyDetectorViewerSettings(self, settingsFile, surfNum=None,
                                      detectNum=None, showAs=None, rowcolNum=None, 
                                      zPlaneNum=None, scale=None, smooth=None, 
                                      dType=None, zrd=None, dfilter=None, 
                                      maxPltScale=None, minPltScale=None, 
                                      outFileName=None):
        """Modify an existing detector viewer settings (configuration) file 

        Only those parameters that are non-None or non-zero-length (in
        case of tuples) will be set.

        Parameters
        ----------
        settingsFile : string
            filename of the settings file including path and extension
        surfNum : integer, optional
            the surface number. Use 1 for Non-Sequential mode
        detectNum : integer, optional
            the detector number
        showAs : integer, optional
            0 = full pixel data; 1 = cross section row; 2 = cross 
            section column. For Graphics Windows see Notes below. 
        rowcolNum: integer, optional
            the row or column number for cross section plots
        zPlaneNum : integer, optional
            the Z-Plane number for detector volumes
        scale : integer, optional 
            the scale mode. Use 0 for linear, 1 for Log -5, 2 for Log -10, and
            3 for Log - 15.
        smooth : integer, optional 
            the smoothing value 
        dType : integer, optional 
            use 0 for incoherent irradiance, 1 for coherent irradiance, 2 for
            coherent phase, 3 for radiant intensity, 4 for radiance (position 
            space), and 5 for radiance (angle space).
        zrd : string, optional 
            the ray data base name, or null for none.
        dfilter : string, optional 
            the filter string  
        maxPltScale : float, optional 
            the maximum plot scale 
        minPltScale : float, optional 
            the minimim plot scale 
        outFileName : string, optional
            the output file name 

        Returns
        -------
        statusTuple : tuple or -1
            tuple of codes returned by ``zModifySettings()`` for each
            non-None parameters. The status codes are as follows:
            0 = no error;
            -1 = invalid file;
            -2 = incorrect version number;
            -3 = file access conflict

            The function returns -1 if ``settingsFile`` is invalid.

        See Also
        --------
        zSetDetectorViewerSettings(), zGetDetectorViewer()
        """
        sTuple = [] # status tuple
        if (_os.path.isfile(settingsFile) and
            settingsFile.lower().endswith('.cfg')):
            dst = settingsFile
        else:
            return -1
        if surfNum is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_SURFACE", surfNum))
        if detectNum is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_DETECTOR", detectNum))
        if showAs is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_SHOW", showAs))
        if rowcolNum is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_ROWCOL", rowcolNum))
        if zPlaneNum is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_ZPLANE", zPlaneNum))
        if scale is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_SCALE", scale))
        if smooth is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_SMOOTHING", smooth))
        if dType is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_DATA", dType))
        if zrd is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_ZRD", zrd))
        if dfilter is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_FILTER", dfilter))
        if maxPltScale is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_MAXPLOT", maxPltScale))
        if minPltScale is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_MINPLOT", minPltScale))
        if outFileName is not None:
            sTuple.append(self.zModifySettings(dst, "DVW_OUTPUTFILE", outFileName))
        return tuple(sTuple)

    def zSetDetectorViewerSettings(self, settingsFile=None, surfNum=None, 
                                   detectNum=None, showAs=None, rowcolNum=None, 
                                   zPlaneNum=None, scale=None, smooth=None, 
                                   dType=None, zrd=None, dfilter=None, 
                                   maxPltScale=None, minPltScale=None, 
                                   outFileName=None):
        """Create and set a new detector viewer settings file starting
        from the "reset" settings state of the most basic lens in Zemax 

        To modify an existing detector viewer settings file, use 
        ``zModifyDetectorViewerSettings()``. Only those parameters that 
        are non-None will be set 

        Parameters
        ---------- 
         settingsFile : string, optional
            full name of the settings file (with .CFG extension). If ``None``, 
            then a CFG file with the name of the lens followed by the string 
            '_pyzdde_DVW.CFG' will be created in the same directory as the 
            lens file and returned         
        surfNum : integer, optional
            the surface number. Use 1 for Non-Sequential mode
        detectNum : integer, optional
            the detector number
        showAs : integer, optional
            0 = full pixel data; 1 = cross section row; 2 = cross 
            section column. For Graphics Windows see Notes below. 
        rowcolNum: integer, optional
            the row or column number for cross section plots
        zPlaneNum : integer, optional
            the Z-Plane number for detector volumes
        scale : integer, optional 
            the scale mode. Use 0 for linear, 1 for Log -5, 2 for Log -10, and
            3 for Log - 15.
        smooth : integer, optional 
            the smoothing value 
        dType : integer, optional 
            use 0 for incoherent irradiance, 1 for coherent irradiance, 2 for
            coherent phase, 3 for radiant intensity, 4 for radiance (position 
            space), and 5 for radiance (angle space).
        zrd : string, optional 
            the ray data base name, or null for none.
        dfilter : string, optional 
            the filter string  
        maxPltScale : float, optional 
            the maximum plot scale 
        minPltScale : float, optional 
            the minimim plot scale 
        outFileName : string, optional
            the output file name 

        Returns
        -------
        settingsFile : string
            the full name, including path and extension, of the just
            created settings file

        Notes
        ----- 
        The meaning of the integer value of ``showAs`` depends upon the type 
        of window displayed -- For Graphics Windows, use 0 for grey scale, 
        1 for inverted grey scale, 2 for false color, 3 for inverted 
        false color, 4 for cross section row, and 5 for cross section 
        column; For Text Windows (which is mostly likely the case when 
        using externally), use 0 for full pixel data, 1 for cross 

        See Also
        -------- 
        zGetDetectorViewer(), zModifyDetectorViewerSettings()
        """
        clean_cfg = 'RESET_SETTINGS_DVW.CFG'
        src = _os.path.join(_pDir, 'ZMXFILES', clean_cfg)
        if settingsFile:
            dst = settingsFile
        else:
            filename_partial = _os.path.splitext(self.zGetFile())[0]
            dst =  filename_partial + '_pyzdde_DVW.CFG'
            self._filesCreated.add(dst)
        try:
            _shutil.copy(src, dst)
        except IOError:
            print("ERROR: Invalid settingsFile {}".format(dst))
            return
        else:
            self.zModifyDetectorViewerSettings(dst, surfNum, detectNum, showAs, 
                rowcolNum, zPlaneNum, scale, smooth, dType, zrd, dfilter, maxPltScale,
                minPltScale, outFileName)   
            return dst        


    # Aberration coefficients analysis functions
    def zGetSeidelAberration(self, which='wave', txtFile=None, keepFile=False):
        """Return the Seidel Aberration coefficients

        Parameters
        ----------
        which : string, optional
            'wave' = Wavefront aberration coefficient (summary) is
                     returned;
            'aber' = Seidel aberration coefficients (total) is returned
            'both' = both Wavefront (summary) and Seidel aberration
                     (total) coefficients are returned
        txtFile : string, optional
            if passed, the seidel text file will be named such. Pass a
            specific txtFile if you want to dump the file into a separate
            directory.
        keepFile : bool, optional
            if ``False`` (default), the Seidel text file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the Seidel text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are availabl

        Returns
        -------
        sac : dictionary or tuple (see below)
            - if 'which' is 'wave', then a dictionary of Wavefront
              aberration coefficient summary is returned;
            - if 'which' is 'aber', then a dictionary of Seidel total
              aberration coefficient is returned;
            - if 'which' is 'both', then a tuple of dictionaries containing
              Wavefront aberration coefficients and Seidel aberration
              coefficients is returned.
        """
        settings = _txtAndSettingsToUse(self, txtFile, 'None', 'Sei')
        textFileName, _, _ = settings
        ret = self.zGetTextFile(textFileName,'Sei', 'None', 0)
        assert ret == 0
        recSystemData = self.zGetSystem() # Get the current system parameters
        numSurf = recSystemData[0]
        line_list = _readLinesFromFile(_openFile(textFileName))
        seidelAberrationCoefficients = {}         # Aberration Coefficients
        seidelWaveAberrationCoefficients = {}     # Wavefront Aberr. Coefficients
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

    def zGetZernike(self, which='fringe', settingsFile=None, txtFile=None,
                     keepFile=False, timeout=5):
        """returns the Zernike Fringe, Standard, or Annular coefficients
        for the currently loaded lens file.

        It provides similar functionality to ZPL command "GETZERNIKE". The
        only difference is that this function returns "Peak to valley to
        centroid" in the `zInfo` metadata instead of "RMS to the zero OPD
        line)

        Parameters
        ----------
        which : string, optional
            ``fringe`` for "Fringe" zernike terms (default), ``standard``
            for "Standard" zernike terms, and ``annular`` for "Annular"
            zernike terms.
        settingsFile : string, optional
            * if passed, the aberration coefficient analysis will be called
              with the given configuration file (settings);
            * if no ``settingsFile`` is passed, and a config file ending
              with the same name as the lens-file post-fixed with
              "_pyzdde_ZFR.CFG"/"_pyzdde_ZST.CFG"/"_pyzdde_ZAT.CFG" is
              present, the settings from this file will be used;
            * if no ``settingsFile`` and no file-name post-fixed with
              "_pyzdde_ZFR.CFG"/"_pyzdde_ZST.CFG"/"_pyzdde_ZAT.CFG" is
              found, but a config file with the same name as the lens file
              is present, the settings from that file will be used;
            * if none of the above types of settings file is found, then a
              default settings will be used
        txtFile : string, optional
            if passed, the aberration coefficient analysis text file will
            be named such. Pass a specific ``txtFile`` if you want to dump
            the file into a separate directory.
        keepFile : bool, optional
            if ``False`` (default), the aberration coefficient text file
            will be deleted after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the analysis text file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)
        timeout : integer, optional
            timeout in seconds.

        Returns
        -------
        zInfo : named tuple
            the 8-tuple contains 1. Peak to Valley (to chief), 2. Peak to
            valley (to centroid), 3. RMS to chief ray, 4. RMS to image
            centroid, 5. Variance, 6. Strehl ratio, 7. RMS fit error, and
            8. Maximum fit error. All  parameters except for Strehl ratio
            has units of waves.
        zCoeff : 1-D named tuple
            the actual Zernike Fringe, Standard, or Annular coefficients.
            The coefficient names conform to the Zemax manual naming
            staring from Z1, Z2, Z3 .... (see example below)

        Notes
        -----
        1. As of current writing, Zemax doesn't provide a way to modify the
           parameters of any aberration coefficient settings file through
           extensions. Thus a settings file for the aberration coefficients
           analysis has to be created manually using the Zemax menu (as
           opposed to programmatic creation and modification of settings)

        Examples
        --------
        >>> zInfo, zCoeff = ln.zGetZernike(which='fringe')
        >>> zInfo
        zInfo(pToVChief=0.08397624, pToVCentroid=0.08397624, rmsToChief=0.02455132, rmsToCentroid=0.02455132, variance=0.00060277, strehl=0.9764846, rmsFitErr=0.0, maxFitErr=0.0)
        >>> print(zInfo.rmsToChief)
        0.02455132
        >>> print(zCoeff)
        zCoeff(Z1=-0.55311265, Z2=0.0, Z3=0.0, Z4=-0.34152763, Z5=0.0, Z6=0.0, Z7=0.0, Z8=0.0, Z9=0.19277286, Z10=0.0, Z11=0.0, Z12=0.0, Z13=0.0, Z14=0.0, Z15=0.0, Z16=-0.01968138, Z17=0.0, Z18=0.0, Z19=0.0, Z20=0.0, Z21=0.0, Z22=0.0, Z23=0.0, Z24=0.0, Z25=-0.00091852, Z26=0.0, Z27=0.0, Z28=0.0, Z29=0.0, Z30=0.0, Z31=0.0, Z32=0.0, Z33=0.0, Z34=0.0, Z35=0.0, Z36=-3.368e-05, Z37=-1.44e-06)
        >>> print(zCoeff.Z1) # zCoeff.Z1 is same as zCoeff[0]
        -0.55311265
        """
        anaTypeDict = {'fringe':'Zfr', 'standard':'Zst', 'annular':'Zat'}
        assert which in anaTypeDict
        anaType = anaTypeDict[which]
        settings = _txtAndSettingsToUse(self, txtFile, settingsFile, anaType)
        textFileName, cfgFile, getTextFlag = settings
        ret = self.zGetTextFile(textFileName, anaType, cfgFile, getTextFlag,
                                timeout)
        assert ret == 0
        line_list = _readLinesFromFile(_openFile(textFileName))
        line_list_len = len(line_list)

        # Extract Meta data
        meta_patterns = ["Peak to Valley\s+\(to chief\)",
                         "Peak to Valley\s+\(to centroid\)",
                         "RMS\s+\(to chief\)",
                         "RMS\s+\(to centroid\)",
                         "Variance",
                         "Strehl Ratio",
                         "RMS fit error",
                         "Maximum fit error"]
        meta = []
        for i, pat in enumerate(meta_patterns):
            meta_line = line_list[_getFirstLineOfInterest(line_list, pat)]
            meta.append(float(_re.search(r'\d{1,3}\.\d{4,8}', meta_line).group()))

        info = _co.namedtuple('zInfo', ['pToVChief', 'pToVCentroid', 'rmsToChief',
                                        'rmsToCentroid', 'variance', 'strehl',
                                        'rmsFitErr', 'maxFitErr'])
        zInfo = info(*meta)

        # Extract coefficients
        start_line_pat = "Z\s+1\s+-?\d{1,3}\.\d{4,8}"
        start_line = _getFirstLineOfInterest(line_list, start_line_pat)
        coeff_pat = _re.compile("-?\d{1,3}\.\d{4,8}")
        zCoeffs = [0]*(line_list_len - start_line)

        for i, line in enumerate(line_list[start_line:]):
            zCoeffs[i] = float(_re.findall(coeff_pat, line)[0])

        zCoeffId = _co.namedtuple('zCoeff',
                    ['Z{}'.format(i+1) for i in range(line_list_len - start_line)])
        zCoeff = zCoeffId(*zCoeffs)

        if not keepFile:
            _deleteFile(textFileName)
        return zInfo, zCoeff

    # -------------------
    # Tools functions
    # -------------------
    
    # System modification functions
    def zLensScale(self, factor=2.0, ignoreSurfaces=None):
        """Scale the lens design by factor specified.

        ``Usage: zLensScale([factor,ignoreSurfaces]) -> ret``

        Parameters
        ----------
        factor : float
            the scale factor. If no factor are passed, the design will
            be scaled by a factor of 2.0
        ignoreSurfaces : tuple
            a tuple of surfaces that are not to be scaled. Such as
            (0,2,3) to ignore surfaces 0 (object surface), 2 and 3.
            Or (OBJ, 2, STO, IMG) to ignore object surface, surface
            number 2, stop surface and image surface.

        Returns
        -------
        status : integer
            0 = success; 1 = success with warning; -1 = failure;

        .. warning::

            1. This function implementation is not yet complete.
                * Not all surfaces have been implemented.
                * ``ignoreSurface`` not implemented yet.
            2. (Limitations) Cannot scale pupil shift x,y, and z in the
               General settings as Zemax hasn't provided any command to
               do so using the extensions. The pupil shift values are also
               scaled, when a lens design is scaled, when the ray-aiming
               is on. However, this is not a serious limitation for most
               cases.
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

    # Design functions
    def zOptimize2(self, numCycle=1, algo=0, histLen=5, precision=1e-12,
                   minMF=1e-15, tMinCycles=5, tMaxCycles=None, timeout=None):
        """A wrapper around zOptimize() providing few control features

        Parameters
        ----------
        numCycles : integer
            number of cycles per DDE call to optimization (default=1)
        algo : integer
            0=DLS, 1=Orthogonal descent (default=0)
        histLen : integer
            length of the array of past merit functions returned from each
            DDE call to ``zOptimize()`` for determining steady state of
            merit function values (default=5)
        precision : float
            minimum acceptable absolute difference between the merit-
            function values in the array for steady state computation
            (default=1e-12)
        minMF : float
            minimum Merit Function following which to the optimization
            loop is to be terminated even if a steady state hasn't reached.
            This might be useful if a target merit function is desired.
        tMinCycles : integer
            total number of cycles to run optimization at the very least.
            This is NOT the number of cycles per DDE call, but it is
            calculated by multiplying the number of cycles per DDL
            optimize call to the total number of DDE calls (default=5).
        tMaxCycles : integer
            the maximum number of cycles after which the optimizaiton
            should be terminated even if a steady state hasn't reached
        timeout : integer
            timeout value, in seconds, used in each pass

        Returns
        -------
        finalMerit : float
            the final merit function.
        tCycles : integer
            total number of cycles calculated by multiplying the number
            of cycles per DDL optimize call to the total number of DDE
            calls.

        Notes
        -----
        ``zOptimize2()`` basically calls ``zOptimize()`` mutiple number of
        times in a loop. It can be useful if a large number of optimization
        cycles are required.
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

    # Other functions
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
        If the macro path is different from the default macro path at 
        ``<data>/Macros``, then first use ``zSetMacroPath()`` to set the 
        macropath and then use ``zExecuteZPLMacro()``.

        .. warning::

          1. can only "execute" an existing ZPL macro. i.e. you can't 
             create a ZPL macro on-the-fly and execute it.
          2. If it is required to redirect the result of executing the ZPL 
             to a text file, modify the ZPL macro in the following way:

            -   Add the following two lines at the beginning of the file:
                ``CLOSEWINDOW`` # to suppress the display of default text window
                ``OUTPUT "full_path_with_extension_of_result_fileName"``
            -   Add the following line at the end of the file:
                ``OUTPUT SCREEN`` # close the file and re-enable screen printing

          3. If there are more than one macros which have the same first 3 letters
             then the top macro in the list as sorted by the filesystem 
             will be executed.
        """
        status = -1
        if self._macroPath:
            zplMpath = self._macroPath
        else:
            zplMpath = _os.path.join(self.zGetPath()[0], 'Macros')
        macroList = [f for f in _os.listdir(zplMpath)
                     if f.endswith(('.zpl','.ZPL')) and f.startswith(zplMacroCode)]
        if macroList:
            zplCode = macroList[0][:3]
            status = self.zOpenWindow(zplCode, True, timeout)
        return status

    def zSetMacroPath(self, macroFolderPath):
        """Set the full path name to the macro folder

        Parameters
        ----------
        macroFolderPath : string
            full-path name of the macro folder path. Also, this folder
            path should match the folder path specified for Macros in the
            Zemax Preferences setting.

        Returns
        -------
        status : integer
            0 = success; -1 = failure

        Notes
        -----
        Use this method to set the full-path name of the macro folder
        path if it is different from the default path at <data>/Macros

        See Also
        --------
        zExecuteZPLMacro()
        """
        if _os.path.isabs(macroFolderPath):
            self._macroPath = macroFolderPath
            return 0
        else:
            return -1

# -------------------
# Report functions
# -------------------
    def zGetImageSpaceNA(self):
        """Return the Image Space Numerical Aperture (ISNA) of the lens

        Parameters
        ----------
        None

        Returns
        -------
        isna : real
            image space numerical aperture

        Notes
        -----
        1. The ISNA is calculated using paraxial ray tracing. It is defined
           as the index of the image space multiplied by the sine of the
           angle between the paraxial on-axis chief ray and the paraxial
           on-axis +y marginal ray calculated at the defined conjugates for
           the primary wavelength [UPRT]_.
        2. Relation to F-number :
           ``isna = pyz.fnum2numAper(paraxial_working_fnumber)``

        References
        ----------
        .. [UPRT] Understanding Paraxial Ray-Tracing, Mark Nicholson, Zemax
                  Knowledgebase, July 21, 2005.

        See Also
        -------- 
        pyz.numAper2fnum()
        """
        prim_wave_num = self.zGetPrimaryWave()
        last_surf = self.zGetNumSurf()
        # Trace paraxial on-axis chief ray at primary wavelength
        chief_ray_dat = self.zGetTrace(prim_wave_num, mode=1, surf=last_surf,
                                       hx=0, hy=0, px=0, py=0)
        chief_angle = _math.asin(chief_ray_dat[6])
        # Trace paraxial marginal ray at primary wavelength
        margi_ray_dat = self.zGetTrace(prim_wave_num, mode=1, surf=last_surf,
                                       hx=0, hy=0, px=0, py=1)
        margi_angle = _math.asin(margi_ray_dat[6])
        index = self.zGetIndexPrimWave(last_surf)
        return index*_math.sin(chief_angle - margi_angle)

    def zGetIndexPrimWave(self, surfNum):
        """Returns the index of refraction at primary wavelength for the
        specified surface

        Emulates the ZPL macro ``INDX(surface)``

        Parameters
        ----------
        surfNum : integer
            surface number

        Returns
        -------
        index : float
            index of refraction at primary wavelength

        See Also
        --------
        zGetIndex()
        """
        prime_wave_num = self.zGetPrimaryWave()
        return self.zGetIndex(surfNum)[prime_wave_num-1]


    def zGetHiatus(self, txtFile=None, keepFile=False):
        """Returns the Hiatus, which is the distance between the two
        principal planes of the optical system

        Parameters
        ----------
        txtFile : string, optional
            if passed, the prescription file will be named such. Pass a
            specific ``txtFile`` if you want to dump the file into a
            separate directory.
        keepFile : bool, optional
            if ``False`` (default), the prescription file will be deleted
            after use.
            If ``True``, the file will persist. If ``keepFile`` is ``True``
            but a ``txtFile`` is not passed, the prescription file will be
            saved in the same directory as the lens (provided the required
            folder access permissions are available)

        Returns
        -------
        hiatus : float
            the value of the hiatus

        Notes
        -----
        The hiatus is also known as the Null space or nodal space or the
        interstitium.
        """
        settings = _txtAndSettingsToUse(self, txtFile, 'None', 'Pre')
        textFileName, _, _ = settings

        sysProp = self.zGetSystem()
        numSurf = sysProp.numSurf

        # Since the object space cardinal points are reported w.r.t. the
        # surface 1, ensure that surface 1 is global reference surface
        if sysProp.globalRefSurf is not 1:
            self.zSetSystem(unitCode=sysProp.unitCode, stopSurf=sysProp.stopSurf,
                            rayAimingType=sysProp.rayAimingType, temp=sysProp.temp,
                            pressure=sysProp.pressure, globalRefSurf=1)

        ret = self.zGetTextFile(textFileName, 'Pre', "None", 0)
        assert ret == 0
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
            # Calculate the hiatus (only if count > 0) as
            hiatus = abs(ima_z + principalPlane_imgSpace - principalPlane_objSpace)

        # Restore the Global ref surface if it was changed
        if sysProp.globalRefSurf is not 1:
            self.zSetSystem(unitCode=sysProp.unitCode, stopSurf=sysProp.stopSurf,
                            rayAimingType=sysProp.rayAimingType, temp=sysProp.temp,
                            pressure=sysProp.pressure,
                            globalRefSurf=sysProp.globalRefSurf)
        if not keepFile:
            _deleteFile(textFileName)
        return hiatus

    def zGetPupilMagnification(self):
        """Return the pupil magnification, which is the ratio of the
        exit-pupil diameter to the entrance pupil diameter.

        The pupils are paraxial pupils. 

        Parameters
        ----------
        None

        Returns
        -------
        pupilMag : real
            the pupil magnification
        """
        _, _, ENPD, ENPP, EXPD, EXPP, _, _ = self.zGetPupil()
        return (EXPD/ENPD)

    def zGetOpticalPathLength(self, surf1=0, surf2=2, hx=0, hy=0, px=0, py=0):
        """Returns the total optical path length (OPL) between surfaces
        surf1 and surf2 for a ray traced at primary wavelength

        Parameters
        ----------
        surf1 : integer, optional
            start surface number
        surf2 : integer, optional
            end surface number
        hx : float, optional
            normalized field coordinate along x
        hy : float, optional
            normalized field coordinate along y
        px : float, optional
            normalized pupil coordinate along x
        py : float, optional
            normalized pupil coordinate along y

        Returns
        -------
        oplen : float
            total optical path length (including refraction and phase
            surfaces) between surfaces

        Notes
        -----
        The function uses the optimization operand "PLEN" to retrieve
        the value of the optical path length

        See Also
        --------
        zGetOpticalPathDifference()
        """
        oplen = self.zOperandValue('PLEN', surf1, surf2, hx, hy, px, py)
        return oplen

    def zGetOpticalPathDifference(self, hx=0, hy=0, px=0, py=0, ref=0, wave=None):
        """Returns the optical path difference (OPD) with respect to the
        chief ray or mean OPD in waves

        Parameters
        ----------
        hx : float, optional
            normalized field coordinate along x
        hy : float, optional
            normalized field coordinate along y
        px : float, optional
            normalized pupil coordinate along x
        py : float, optional
            normalized pupil coordinate along y
        ref : integer, optional
            integer code to indicate reference ray/OPD.
            0 = chief ray (Default); 1 = mean OPD over the pupil;
            2 = mean OPD over the pupil with tilt removed
        wave : integer, optional
            wavelength number to trace ray. If ``None``, the ray is
            traced at the primary wavelength.

        Returns
        -------
        opd : float
            optical path difference

        See Also
        --------
        zGetOpticalPathLength()
        """
        if ref == 2:
            code = 'OPDX'
        elif ref == 1:
            code = 'OPDM'
        elif ref == 0:
            code = 'OPDC'
        else:
            raise ValueError("Unexpected ref input value")
        if wave is None:
            wave = self.zGetWave(self.zGetPrimaryWave()).wavelength
        opd = self.zOperandValue(code, 0, wave, hx, hy, px, py)
        return opd

    def zGetSemiDiameter(self, surfNum):
        """Get the Semi-Diameter value of the surface with number `surfNum`

        Parameters
        ---------- 
        surfNum : integer 
            surface number 

        Returns
        ------- 
        semidia : real 
            semi-diameter of the surface 
        """
        return self.zGetSurfaceData(surfNum=surfNum, code=self.SDAT_SEMIDIA)

    def zSetSemiDiameter(self, surfNum, value=0):
        """Set the Semi-Diameter of the surface with number `surfNum`.  
        
        A "fixed" solve type is set on the semi-diameter of the surface.
    
        Parameters
        ---------- 
        surfNum : integer
            surface number
        value : real, optional
            value of the semi-diameter to set
    
        Returns
        ------- 
        semidia : real
            value of the semi-diameter of the surface after setting it.
        """
        self.zSetSolve(surfNum, self.SOLVE_SPAR_SEMIDIA, self.SOLVE_SEMIDIA_FIXED)
        return self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_SEMIDIA, value=value)

    def zGetThickness(self, surfNum):
        """Get the Thickness value of the surface with number `surfNum`

        Parameters
        ---------- 
        surfNum : integer 
            surface number 

        Returns
        ------- 
        thick : real 
            thickness of the surface 
        """
        return self.zGetSurfaceData(surfNum=surfNum, code=self.SDAT_THICK)

    def zSetThickness(self, surfNum, value=0):
        """Set the thickness of the surface with number `surfNum`.

        Parameters
        ---------- 
        surfNum: integer
            surface number 
        value : real, optional
            value of the thickness to set 

        Returns
        ------- 
        thick : real 
            value of the thickness of the surface after setting it.
        """
        return self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_THICK, value=value)

    def zGetRadius(self, surfNum):
        """Get the radius of the surface with number `surfNum`.

        Parameters
        ---------- 
        surfNum : integer 
            surface number 

        Returns
        ------- 
        radius : real 
            radius of the surface 
        """
        value = self.zGetSurfaceData(surfNum=surfNum, code=self.SDAT_CURV)
        radius = 1.0/value if value else 1E10
        return radius

    def zSetRadius(self, surfNum, value=1E10):
        """Set the radius of the surface with number `surfNum`.

        Parameters
        ---------- 
        surfNum : integer 
            surface number
        value : real 
            radius of the surface  

        Returns
        ------- 
        radius : real 
            radius of the surface 
        """
        curv = 1.0/value if value else 1E10
        ret = self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_CURV, value=curv)
        return 1.0/ret if ret else 1E10  

    def zSetGlass(self, surfNum, value=''):
        """Set the glass of the surface with number `surfNum`

        Parameters
        ---------- 
        surfNum : integer 
            surface number 
        value : string 
            valid glass string code 

        Returns
        ------- 
        glass : string 
            glass for the surface 
        """
        return self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_GLASS, value=value)

    def zGetConic(self, surfNum):
        """Get the conic value of the surface with number `surfNum`

        Parameters
        ---------- 
        surfNum : integer 
            surface number 

        Returns
        ------- 
        conic : real 
            conic of the surface
        """
        return self.zGetSurfaceData(surfNum=surfNum, code=self.SDAT_CONIC)

    def zSetConic(self, surfNum, value=0):
        """Set the conic value of the surface with number `surfNum`

        Parameters
        ---------- 
        surfNum : integer 
            surface number 
        value : real 
            conic value

        Returns
        ------- 
        conic : real 
            conic of the surface
        """
        return self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_CONIC, value=value)
        
    def zInsertDummySurface(self, surfNum, comment='dummy', thick=None, semidia=None):
        """Insert dummy surface at surface number indicated by `surfNum` 
    
        Parameters
        ---------- 
        surfNum : integer
            surface number at which to insert the dummy surface 
        comment : string, optional, default is 'dummy'
            comment on the surface 
        thick : real, optional
            thickness of the surface 
        semidia : real, optional 
            semi diameter of the surface 
            
        Returns
        -------
        nsur : integer
            total number of surfaces in the LDE including the new dummy surface.
        """
        self.zInsertSurface(surfNum)
        self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_COMMENT, value=comment)
        if thick is not None:
            self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_THICK, value=thick)
        if semidia is not None:
            self.zSetSemiDiameter(surfNum=surfNum, value=semidia)
        return self.zGetNumSurf()
        
    def zInsertCoordinateBreak(self, surfNum, xdec=0.0, ydec=0.0, xtilt=0.0, ytilt=0.0,
                               ztilt=0.0, order=0, thick=None, comment=None):
        """Insert Coordinate Break at the surface position indicated by `surfNum`
        
        Parameters
        ----------
        surfNum : integer
            surface number at which to insert the coordinate break 
        xdec : float, optional, default = 0.0
            decenter x (in lens units)
        ydec : float, optional, default = 0.0
            decenter y (in lens units)
        xtilt : float, optional, default = 0.0
            tilt about x (degrees)
        ytilt : float, optional, default = 0.0
            tilt about y (degrees)
        ztilt : float, optional, default = 0.0
            tilt about z (degrees)
        order : integer (0/1), optional, default = 0
            0 = decenter then tilt; 1 = tilt then decenter
        thick : real, optional
            set the thickness of the cb surface 
        comment : string, optional 
            surface comment 

        Returns
        ------- 
        ret : integer
            0 if no error   
        """
        self.zInsertSurface(surfNum=surfNum) 
        self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_TYPE, value='COORDBRK')
        # set the decenter and tilt values and order 
        params = range(1, 7)
        values = [xdec, ydec, xtilt, ytilt, ztilt, order]
        for par, val in zip(params, values):
            self.zSetSurfaceParameter(surfNum=surfNum, param=par, value=val)
        if thick is not None:
            self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_THICK, value=thick)
        if comment is not None:
             self.zSetSurfaceData(surfNum=surfNum, code=self.SDAT_COMMENT, value=comment)
        return 0
        
        
    def zTiltDecenterElements(self, firstSurf, lastSurf, xdec=0.0, ydec=0.0, xtilt=0.0, 
                              ytilt=0.0, ztilt=0.0, order=0, cbComment1=None, 
                              cbComment2=None, dummySemiDiaToZero=False):
        '''Tilt decenter elements using CBs around the `firstSurf` and `lastSurf`. 
        
        Parameters
        ----------
        firstSurf : integer
            first surface
        lastSurf : integer
            last surface
        xdec : float
            decenter x (in lens units)
        ydec : float
            decenter y (in lens units)
        xtilt : float
            tilt about x (degrees)
        ytilt : float
            tilt about y (degrees)
        ztilt : float
            tilt about z (degrees)
        order : integer (0/1), optional, default = 0
            0 = decenter then tilt; 1 = tilt then decenter
        comment1 : string, optional, default = 'Element Tilt'
            comment on the first CB surface
        comment2 : string, optional, default = 'Element Tilt:return'
            comment on the second CB surface. 
        dummySemiDiaToZero : bool, optional, default = False
            if `True` the semi-diameter of the dummy surface (afer CB2) is 
            set to zero.
        
        Returns
        -------
        cb1 : integer
            surface number of the first coordinate break surface
        cb2 : integer
            surface number of the second (for restoring axis) coordinate 
            break surface
        dummy : integer
            surface number of the dummy surface. It is always `CB2` + 1
        
        Notes
        -----
        1. In total, 3 more surfaces are added to the existing system -- the 
           first CB, before `firstSurf`, the second CB after `lastSurf`, and 
           a dummy surface after the second CB. 
        '''
        numSurfBetweenCBs = lastSurf - firstSurf + 1
        cb1 = firstSurf
        cb2 = cb1 + numSurfBetweenCBs + 1
        dummy = cb2 + 1
        # store the thickness and solve on thickness (if any) of the last surface 
        thick = self.zGetSurfaceData(surfNum=lastSurf, code=self.SDAT_THICK)
        solve = self.zGetSolve(surfNum=lastSurf, code=self.SOLVE_SPAR_THICK)
        # insert surfaces
        self.zInsertSurface(surfNum=cb1) # 1st cb
        self.zInsertSurface(surfNum=cb2) # 2nd cb to restore the original axis
        self.zInsertSurface(surfNum=dummy) # dummy after 2nd cb
        cbComment1 = cbComment1 if cbComment1 else 'Element Tilt'
        self.zSetSurfaceData(surfNum=cb1, code=self.SDAT_COMMENT, value=cbComment1)
        self.zSetSurfaceData(surfNum=cb1, code=self.SDAT_TYPE, value='COORDBRK')
        cbComment2 = cbComment2 if cbComment2 else 'Element Tilt:return'
        self.zSetSurfaceData(surfNum=cb2, code=self.SDAT_COMMENT, value=cbComment2)
        self.zSetSurfaceData(surfNum=cb2, code=self.SDAT_TYPE, value='COORDBRK')
        self.zSetSurfaceData(surfNum=dummy, code=self.SDAT_COMMENT, value='Dummy')
        if dummySemiDiaToZero:
            self.zSetSemiDiameter(surfNum=dummy, value=0)
        # transfer thickness of the surface just before the cb2 (originally 
        # lastSurf) to the dummy surface
        lastSurf += 1  # last surface number incremented by 1 bcoz of cb 1
        self.zSetSurfaceData(surfNum=lastSurf, code=self.SDAT_THICK, value=0.0)
        self.zSetSolve(lastSurf, self.SOLVE_SPAR_THICK, self.SOLVE_THICK_FIXED)
        self.zSetSurfaceData(surfNum=dummy, code=self.SDAT_THICK, value=thick)
        # transfer the solve on the thickness (if any) of the surface just before
        # the cb2 (originally lastSurf) to the dummy surface. The param1 of 
        # solve type "Thickness" may need to be modified before transferring.
        if solve[0] in {5, 7, 8, 9}: # param1 is a integer surface number
            param1 = int(solve[1]) if solve[1] < cb1 else int(solve[1]) + 1
        else: # param1 is a floating value, or macro name
            param1 = solve[1]
        self.zSetSolve(dummy, self.SOLVE_SPAR_THICK, solve[0], param1, solve[2], 
                       solve[3], solve[4])
        # use pick-up solve on glass surface of dummy to pickup from lastSurf
        self.zSetSolve(dummy, self.SOLVE_SPAR_GLASS, self.SOLVE_GLASS_PICKUP, lastSurf)
        # use pick-up solves on second CB; set scale factor of -1 to lock the second
        # cb to the first.
        pickupcolumns = range(6, 11)
        params = [self.SOLVE_SPAR_PAR1, self.SOLVE_SPAR_PAR2, 
                  self.SOLVE_SPAR_PAR3, self.SOLVE_SPAR_PAR4, self.SOLVE_SPAR_PAR5]
        offset, scale = 0, -1
        for para, pcol in zip(params, pickupcolumns):
            self.zSetSolve(cb2, para, self.SOLVE_PARn_PICKUP, cb1, offset, scale, pcol)       
        # Set solves to co-locate the two CBs
        # use position solve to track back through the lens
        self.zSetSolve(lastSurf, self.SOLVE_SPAR_THICK, self.SOLVE_THICK_POS, cb1 , 0)
        # use a pickup solve to restore position at the back of the lastSurf
        self.zSetSolve(cb2, self.SOLVE_SPAR_THICK, self.SOLVE_THICK_PICKUP, 
                       lastSurf, scale, offset, 0)
        # set the appropriate orders on the surfaces
        if order: # Tilt and then decenter
            cb1Ord, cb2Ord = 1, 0
        else: # Decenter and then tilt (default)
            cb1Ord, cb2Ord = 0, 1
        self.zSetSurfaceParameter(surfNum=cb1, param=6, value=cb1Ord)    
        self.zSetSurfaceParameter(surfNum=cb2, param=6, value=cb2Ord)
        # set the decenter and tilt values in the first cb
        params = range(1, 6)
        values = [xdec, ydec, xtilt, ytilt, ztilt]
        for par, val in zip(params, values):
            self.zSetSurfaceParameter(surfNum=cb1, param=par, value=val)
        self.zGetUpdate()
        return cb1, cb2, dummy

    def zInsertNSCSourceEllipse(self, surfNum=1, objNum=1, x=0.0, y=0.0, z=0.0, 
                                tiltX=0.0, tiltY=0.0, tiltZ=0.0, xHalfWidth=0, 
                                yHalfWidth=0, numLayRays=20, numAnaRays=500,
                                refObjNum=0, insideOf=0, power=1, waveNum=0,
                                srcDist=0.0, cosExp=0.0, gaussGx=0.0, gaussGy=0.0, 
                                srcX=0.0, srcY=0.0, minXHalfWidth=0.0, minYHalfWidth=0.0,
                                color=0, comment='', overwrite=False):
        """Insert a new NSC source ellipse at the location indicated by the 
        parameters ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer, optional
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer, optional
            object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        xHalfWidth, yHalfWidth : floats
            half widths along x and y axis
        numLayRays, numAnaRays : integers, optional
            number of layout rays and analysis rays respectively
        refObjNum : integer, optional
            reference object number
        insideOf : integer, optional
            inside of object number
        power : float, optional
            power in Watts 
        waveNum : integer, optional
            the wave number
        srcDist, cosExp, gaussGx, gaussGy, srcX, srcY, minXHalfWidth, minYHalfWidth : floats
            see the manual for details
        color : integer, optional
            The pen color to use when drawing rays from this source. If 0, 
            the default color will be chosen.
        comment : string, optional
            comment for the object
        overwrite : bool, optional
            if `False` (default), a new object is inserted at the position and existing
            objects (if any) are pushed. If `True`, then existing at the ``objNum`` is
            overwritten

        Returns
        -------
        None
        
        Note
        ----
        If an object with the same number as ``objNum`` already exist in the NSCE, 
        that (and any subsequent) object is pushed by one row, unless ``overwrite`` 
        is ``True``.
        
        See Also
        --------
        zInsertNSCSourceRectangle()
        """
        numObjsExist = self.zGetNSCData(surfNum, code=0)
        if objNum > numObjsExist + 1:
            raise ValueError('objNum ({}) cannot be greater than {}.'
            .format(objNum, numObjsExist+1))
        if not overwrite:
            assert self.zInsertObject(surfNum, objNum) == 0, \
            'Error inserting object at object Number {}'.format(objNum)
        objData = {0:'NSC_SRCE', 1:comment, 5:refObjNum, 6:insideOf}
        for code, data in objData.iteritems():
            assert self.zSetNSCObjectData(surfNum, objNum, code, data) == data, \
            'Error in setting NSC object code {}'.format(code)
        assert self.zSetNSCPositionTuple(surfNum, objNum, x, y, z, tiltX, tiltY, tiltZ) \
        == (x, y, z, tiltX, tiltY, tiltZ, '')
        param = (numLayRays, numAnaRays, power, waveNum, color, xHalfWidth, yHalfWidth,
                 srcDist, cosExp, gaussGx, gaussGy, srcX, srcY, minXHalfWidth, minYHalfWidth)
        for i, each in enumerate(param, 1):
            assert self.zSetNSCParameter(surfNum, objNum, paramNum=i, data=each) == each, \
            'Error in setting NSC parameter {} to {} at object {}'.format(i, param[i], objNum)                   

    def zInsertNSCSourceRectangle(self, surfNum=1, objNum=1, x=0.0, y=0.0, z=0.0, 
                                tiltX=0.0, tiltY=0.0, tiltZ=0.0, xHalfWidth=0, 
                                yHalfWidth=0, numLayRays=20, numAnaRays=500,
                                refObjNum=0, insideOf=0, power=1, waveNum=0,
                                srcDist=0.0, cosExp=0.0, gaussGx=0.0, gaussGy=0.0, 
                                srcX=0.0, srcY=0.0, color=0, comment='', overwrite=False):
        """Insert a new NSC source rectangle at the location indicated by the 
        parameters ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer, optional
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer, optional
            object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        xHalfWidth, yHalfWidth : floats
            half widths along x and y axis
        numLayRays, numAnaRays : integers, optional
            number of layout rays and analysis rays respectively
        refObjNum : integer, optional
            reference object number
        insideOf : integer, optional
            inside of object number
        power : float, optional
            power in Watts 
        waveNum : integer, optional
            the wave number
        srcDist, cosExp, gaussGx, gaussGy, srcX, srcY : floats
            see the manual for details
        color : integer, optional
            The pen color to use when drawing rays from this source. If 0, 
            the default color will be chosen.
        comment : string, optional
            comment for the object
        overwrite : bool, optional
            if `False` (default), a new object is inserted at the position and existing
            objects (if any) are pushed. If `True`, then existing at the ``objNum`` is
            overwritten

        Returns
        -------
        None
        
        Note
        ----
        If an object with the same number as ``objNum`` already exist in the NSCE, 
        that (and any subsequent) object is pushed by one row, unless ``overwrite`` 
        is ``True``.
        
        See Also
        --------
        zInsertNSCSourceEllipse()
        """
        numObjsExist = self.zGetNSCData(surfNum, code=0)
        if objNum > numObjsExist + 1:
            raise ValueError('objNum ({}) cannot be greater than {}.'
            .format(objNum, numObjsExist+1))
        if not overwrite:
            assert self.zInsertObject(surfNum, objNum) == 0, \
            'Error inserting object at object Number {}'.format(objNum)
        objData = {0:'NSC_SRCR', 1:comment, 5:refObjNum, 6:insideOf}
        for code, data in objData.iteritems():
            assert self.zSetNSCObjectData(surfNum, objNum, code, data) == data, \
            'Error in setting NSC object code {}'.format(code)
        assert self.zSetNSCPositionTuple(surfNum, objNum, x, y, z, tiltX, tiltY, tiltZ) \
        == (x, y, z, tiltX, tiltY, tiltZ, '')
        param = (numLayRays, numAnaRays, power, waveNum, color, xHalfWidth, yHalfWidth,
                 srcDist, cosExp, gaussGx, gaussGy, srcX, srcY)
        for i, each in enumerate(param, 1):
            assert self.zSetNSCParameter(surfNum, objNum, paramNum=i, data=each) == each, \
            'Error in setting NSC parameter {} to {} at object {}'.format(i, param[i], objNum) 
            
    def zInsertNSCEllipse(self, surfNum=1, objNum=1, x=0.0, y=0.0, z=0.0, 
                          tiltX=0.0, tiltY=0.0, tiltZ=0.0, xHalfWidth=0.0,
                          yHalfWidth=0.0, material='', refObjNum=0, insideOf=0, 
                          comment='', overwrite=False):
        """Insert a new NSC ellipse object at the location indicated by the 
        parameters ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer, optional
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer, optional
            object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        xHalfWidth, yHalfWidth : floats
            half widths along x and y axis
        material : string, optional
            material such as ABSORB, MIRROR, etc.
        refObjNum : integer, optional
            reference object number
        insideOf : integer, optional
            inside of object number
        comment : string, optional
            comment for the object
        overwrite : bool, optional
            if `False` (default), a new object is inserted at the position and existing
            objects (if any) are pushed. If `True`, then existing at the ``objNum`` is
            overwritten

        Returns
        -------
        None
        
        Note
        ----
        If an object with the same number as ``objNum`` already exist in the NSCE, 
        that (and any subsequent) object is pushed by one row, unless ``overwrite`` 
        is ``True``.

        See Also
        --------
        zInsertNSCRectangle()
        """
        numObjsExist = self.zGetNSCData(surfNum, code=0)
        if objNum > numObjsExist + 1:
            raise ValueError('objNum ({}) cannot be greater than {}.'
            .format(objNum, numObjsExist+1))
        if not overwrite:
            assert self.zInsertObject(surfNum, objNum) == 0, \
            'Error inserting object at object Number {}'.format(objNum)
        objData = {0:'NSC_ELLI', 1:comment, 5:refObjNum, 6:insideOf}
        for code, data in objData.iteritems():
            assert self.zSetNSCObjectData(surfNum, objNum, code, data) == data, \
            'Error in setting NSC object code {}'.format(code)
        assert self.zSetNSCPositionTuple(surfNum, objNum, x, y, z, tiltX, tiltY, tiltZ, material) \
        == (x, y, z, tiltX, tiltY, tiltZ, material)
        param = (xHalfWidth, yHalfWidth,)
        for i, each in enumerate(param, 1):
            assert self.zSetNSCParameter(surfNum, objNum, paramNum=i, data=each) == each, \
            'Error in setting NSC parameter {} to {} at object {}'.format(i, param[i], objNum) 

    def zInsertNSCRectangle(self, surfNum=1, objNum=1, x=0.0, y=0.0, z=0.0, 
                            tiltX=0.0, tiltY=0.0, tiltZ=0.0, xHalfWidth=0.0,
                            yHalfWidth=0.0, material='', refObjNum=0, insideOf=0, 
                            comment='', overwrite=False):
        """Insert a new NSC rectangle object at the location indicated by the 
        parameters ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer, optional
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer, optional
            object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        xHalfWidth, yHalfWidth : floats
            half widths along x and y axis
        material : string, optional
            material such as ABSORB, MIRROR, etc.
        refObjNum : integer, optional
            reference object number
        insideOf : integer, optional
            inside of object number
        comment : string, optional
            comment for the object
        overwrite : bool, optional
            if `False` (default), a new object is inserted at the position and existing
            objects (if any) are pushed. If `True`, then existing at the ``objNum`` is
            overwritten

        Returns
        -------
        None
        
        Note
        ----
        If an object with the same number as ``objNum`` already exist in the NSCE, 
        that (and any subsequent) object is pushed by one row, unless ``overwrite`` 
        is ``True``.
        
        See Also
        --------
        zInsertNSCEllipse()
        """
        numObjsExist = self.zGetNSCData(surfNum, code=0)
        if objNum > numObjsExist + 1:
            raise ValueError('objNum ({}) cannot be greater than {}.'
            .format(objNum, numObjsExist+1))
        if not overwrite:
            assert self.zInsertObject(surfNum, objNum) == 0, \
            'Error inserting object at object Number {}'.format(objNum)
        objData = {0:'NSC_SRCR', 1:comment, 5:refObjNum, 6:insideOf}
        for code, data in objData.iteritems():
            assert self.zSetNSCObjectData(surfNum, objNum, code, data) == data, \
            'Error in setting NSC object code {}'.format(code)
        assert self.zSetNSCPositionTuple(surfNum, objNum, x, y, z, tiltX, tiltY, tiltZ, material) \
        == (x, y, z, tiltX, tiltY, tiltZ, material)
        param = (xHalfWidth, yHalfWidth,)
        for i, each in enumerate(param, 1):
            assert self.zSetNSCParameter(surfNum, objNum, paramNum=i, data=each) == each, \
            'Error in setting NSC parameter {} to {} at object {}'.format(i, param[i], objNum) 
    
    def zNSCDetectorClear(self, surfNum, detectNum=0):
        """Clear NSC detector data
    
        Parameters
        ----------
        surfNum : integer
            surface number of NSC group (use 1 for pure NSC system)
        detectNum : integer
            the object number of the detector to be cleared. Use 0 to 
            clear all detectors

        Returns
        ------- 
        ret : integer 
            0 if successful
        """
        return self.zNSCDetectorData(surfNum, -detectNum, 0, 0)

    def zInsertNSCDetectorRectangle(self, surfNum=1, objNum=1, x=0.0, y=0.0, z=0.0, 
                                    tiltX=0.0, tiltY=0.0, tiltZ=0.0, xHalfWidth=1.0,
                                    yHalfWidth=1.0, numXPix=1, numYPix=1, material='', 
                                    dType=0, fntOnly=0, refObjNum=0, insideOf=0,
                                    color=0, smooth=0, scale=0, pltScale=0.0,  
                                    psfWaveNum=0, xAngMin=-90.0, xAngMax=90.0, 
                                    yAngMin=-90.0, yAngMax=90.0, pol=0, mirror=0,
                                    comment='', overwrite=False):
        """Insert a new NSC detector rectangle at the location indicated by the 
        parameters ``surfNum`` and ``objNum``

        Parameters
        ----------
        surfNum : integer, optional
            surface number of the NSC group. Use 1 if the program mode is
            Non-Sequential
        objNum : integer, optional
            object number
        x, y, z, tiltX, tiltY, tiltZ : floats, optional
            x, y, z position and tilts about X, Y, and Z axis respectively
        xHalfWidth, yHalfWidth : floats
            half widths along x and y axis
        numXPix, numYPix: integers, optional
            number of pixels along x- and y- axis respectively
        material : string, optional
            material such as ABSORB, MIRROR, etc.
        dType : integer, optional
            whether coherent or incoherent
        fntOnly : integer, optional
            whether detection occurs only on the front surface
        refObjNum : integer, optional
            reference object number
        insideOf : integer, optional
            inside of object number
        color, smooth, scale, pltScale, psfWaveNum :
            see manual for details
        xAngMin, xAngMax, yAngMin, yAngMax, pol, mirror :
            see manual for details
        comment : string, optional
            comment for the object
        overwrite : bool, optional
            if `False` (default), a new object is inserted at the position and existing
            objects (if any) are pushed. If `True`, then existing at the ``objNum`` is
            overwritten

        Returns
        -------
        None
        
        Note
        ----
        If an object with the same number as ``objNum`` already exist in the NSCE, 
        that (and any subsequent) object is pushed by one row, unless ``overwrite`` 
        is ``True``.
        """
        numObjsExist = self.zGetNSCData(surfNum, code=0)
        if objNum > numObjsExist + 1:
            raise ValueError('objNum ({}) cannot be greater than {}.'
            .format(objNum, numObjsExist+1))
        if not overwrite:
            assert self.zInsertObject(surfNum, objNum) == 0, \
            'Error inserting object at object Number {}'.format(objNum)
        objData = {0:'NSC_DETE', 1:comment, 5:refObjNum, 6:insideOf}
        for code, data in objData.iteritems():
            assert self.zSetNSCObjectData(surfNum, objNum, code, data) == data, \
            'Error in setting NSC object code {}'.format(code)
        assert self.zSetNSCPositionTuple(surfNum, objNum, x, y, z, tiltX, tiltY, tiltZ, material) \
        == (x, y, z, tiltX, tiltY, tiltZ, material)
        param = (xHalfWidth, yHalfWidth, numXPix, numYPix,  dType, color, smooth, scale, 
                 pltScale, fntOnly, psfWaveNum, xAngMin, xAngMax, yAngMin, yAngMax, pol, mirror)
        for i, each in enumerate(param, 1):
            assert self.zSetNSCParameter(surfNum, objNum, paramNum=i, data=each) == each, \
            'Error in setting NSC parameter {} to {} at object {}'.format(i, param[i], objNum) 
    
    #%% Interaction friendly (but duplicate) functions
    @property
    def refresh(self):
        """push lens from LDE to DDE server"""
        return self.zGetRefresh()

    @property
    def push(self):
        """push lens from DDE server to LDE and update lens"""
        return self.zPushLens(1) 

    @property
    def update(self):
        """update -- recompute all pupil positions, slovles, etc."""
        return self.zGetUpdate()
    

    #%% IPYTHON NOTEBOOK UTILITY FUNCTIONS

    def ipzCaptureWindowLQ(self, num=1, *args, **kwargs):
        """Capture graphic window from Zemax and display in IPython
        (Low Quality)

        ipzCaptureWindowLQ(num [, *args, **kwargs])-> displayGraphic

        Parameters
        ----------
        num : integer
            the graphic window to capture is indicated by the window
            number ``num``.

        Returns
        -------
        None (embeds image in IPython cell)

        Notes
        -----
        1. This function is useful for quickly capturing a graphic window,
           and embedding into a IPython notebook or QtConsole. The quality
           of JPG image is limited by the JPG export quality from Zemax.
        2. In order to use this function, please copy the ZPL macros from
           PyZDDE\ZPLMacros to the macro directory where Zemax is expecting
           (i.e. as set in Zemax->Preference->Folders)
        3. For earlier versions (before 2010) please use
           ``ipzCaptureWindow()`` for better quality.
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
                      .format(self._macroPath))
                if not self._macroPath:
                    print("Use zSetMacroPath() to set the correct macro path.")
        else:
            print("Couldn't import IPython modules.")

    def ipzCaptureWindow(self, analysisType, percent=12, MFFtNum=0, gamma=0.35,
                         settingsFile=None, flag=0, retArr=False, wait=10):
        """Capture any analysis window from Zemax main window, using
        3-letter analysis code.

        Parameters
        ----------
        analysisType : string
            3-letter button code for the type of analysis
        percent : float
            percentage of the Zemax metafile to display (default=12). Used
            for resizing the large metafile.
        MFFtNum : integer
            type of metafile. 0 = Enhanced Metafile; 1 = Standard Metafile
        gamma : float
            gamma for the PNG image (default = 0.35). Use a gamma value of
            around 0.9 for color surface plots.
        settingsFile : string
            If a valid file name is used for the ``settingsFile``, Zemax
            will use or save the settings used to compute the  metafile,
            depending upon the value of the flag parameter.
        flag : integer
            * 0 = default settings used for the metafile graphic;
            * 1 = settings provided in the settings file, if valid, else
              default settings used;
            * 2 = settings in the settings file, if valid, will be used &
              the settings box for the requested feature will be displayed.
              After the user changes the settings, the graphic will be
              generated using the new settings.
        retArr : boolean
            whether to return the image as an array or not.
            If ``False`` (default), the image is embedded and no array is
            returned;
            If ``True``, an numpy array is returned that may be plotted
            using Matplotlib.
        wait : integer
            time in sec sent to Zemax for the requested analysis to
            complete and produce a file.

        Returns
        -------
        None if ``retArr`` is ``False`` (default). The graphic is embedded
        into the notebook, else ``pixel_array`` (ndarray) if ``retArr``
        is ``True``.

        Notes
        -----
        1. PyZDDE uses ImageMagick's convert program to resize and
           convert the meta file produced by Zemax into a smaller PNG
           file suitable for embedding into IPython cells. A copy of
           convert.exe comes along with PyZDDE. However, the user may
           choose to use a version of ImageMagick that is installed in
           the system already. Please use the module level function
           ``pyz.setImageMagickSettings()`` to do so.
        2. In some environments, Zemax outputs large EMF files [1]_.
           Converting large files may take some time (around 10 sec.)
           to complete.
        3. If the function doesn't work as expected, please check the
           EMF to PNG conversion command being used and the output of
           executing this command by setting the debug print level to
           1 as ``pyz._DEBUG_PRINT_LEVEL=1``, and running the function
           again.
        4. The dataitem `GetMetaFile` has been removed since OpticStudio
           14. Therefore, this function does not work in OpticStudio. 

        References
        ----------
        .. [1] https://github.com/indranilsinharoy/PyZDDE/issues/34

        See Also
        --------
        ipzCaptureWindowLQ():
            low quality screen-shot of a graphic window
        pyz.setImageMagickSettings():
            set ImageMagick settings
        pyz.getImageMagickSettings():
            view current ImageMagick settings
        """
        global _global_IPLoad, _global_mpl_img_load
        global _global_use_installed_imageMagick
        global _global_imageMagick_dir
        if _global_IPLoad:
            tmpImgPath = _os.path.dirname(self.zGetFile())  # dir of the lens file
            if MFFtNum==0:
                ext = 'EMF'
            else:
                ext = 'WMF'
            tmpMetaImgName = "{tip}\\TEMPGPX.{ext}".format(tip=tmpImgPath, ext=ext)
            tmpPngImgName = "{tip}\\TEMPGPX.png".format(tip=tmpImgPath)

            if _global_use_installed_imageMagick:
                cd = _global_imageMagick_dir
            else:
                cd = _os.path.dirname(_os.path.realpath(__file__))
            if MFFtNum==0:
                imagickCmd = ('{cd}\convert.exe \"{MetaImg}\" -flatten '
                              '-resize {per}% -gamma {ga} \"{PngImg}\"'
                              .format(cd=cd,MetaImg=tmpMetaImgName,per=percent,
                                      ga=gamma,PngImg=tmpPngImgName))
            else:
                imagickCmd = ("{cd}\convert.exe \"{MetaImg}\" -resize {per}% \"{PngImg}\""
                              .format(cd=cd,MetaImg=tmpMetaImgName,per=percent,
                                      PngImg=tmpPngImgName))
            _debugPrint(1, "imagickCmd = {}".format(imagickCmd))
            # Get the metafile and display the image
            if not self.zGetMetaFile(tmpMetaImgName,analysisType,
                                     settingsFile,flag):
                if _checkFileExist(tmpMetaImgName, timeout=wait):
                    # Convert Metafile to PNG using ImageMagick's convert
                    startupinfo = _subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= _subprocess.STARTF_USESHOWWINDOW
                    proc = _subprocess.Popen(args=imagickCmd,
                                             stdout=_subprocess.PIPE,
                                             stderr=_subprocess.PIPE,
                                             startupinfo=startupinfo)
                    msg = proc.communicate()
                    _debugPrint(1, "imagickCmd execution return message = "
                                "{}".format(msg))
                    if _checkFileExist(tmpPngImgName, timeout=10):
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
                    print(("Timeout reached before Metafile file was ready. "
                           "This function doesn't work in newer OpticStudio. "
                           "Please consider using ipzCaptureWindowLQ()."))
            else:
                print("Metafile couldn't be created.")
        else:
                print("Couldn't import IPython modules.")

    def ipzGetTextWindow(self, analysisType, sln=0, eln=None, settingsFile=None,
                         flag=0, *args, **kwargs):
        """Print the text output of a Zemax analysis type into a IPython
        cell.

        Parameters
        ----------
        analysisType : string
            3 letter case-sensitive label that indicates the type of the
            analysis to be performed. They are identical to those used for
            the button bar in Zemax. The labels are case sensitive. If no
            label is provided or recognized, a standard raytrace will be
            generated.
        sln : integer, optional
            starting line number (default = 0)
        eln : integer, optional
            ending line number (default = None). If ``None`` all lines in
            the file are printed.
        settingsFile : string, optional
            If a valid file name is used for the ``settingsFile``, Zemax
            will use or save the settings used to compute the text file,
            depending upon the value of the flag parameter.
        flag : integer, optional
            0 = default settings used for the text;
            1 = settings provided in the settings file, if valid, else
                default settings used
            2 = settings provided in the settings file, if valid, will
                be used and the settings box for the requested feature
                will be displayed. After the user makes any changes to
                the settings the text will then be generated using the
                new settings. Please see the ZEMAX manual for details.

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
        """Prints or returns first order data in human readable form

        Parameters
        ----------
        pprint : boolean
            If True (default), the parameters are printed, else a
            dictionary is returned

        Returns
        -------
        firstData : dictionary or None
            if ``pprint`` is True then None
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

    def ipzGetMFE(self, start_row=1, end_row=2, pprint=True):
        """Prints or returns the Oper, Target, Weight and Value parameters
        in the MFE for the specified rows in an IPython notebook cell

        Parameters
        ----------
        start_row : integer, optional
            starting row in the MFE to print (default=1)
        end_row : integer, optional
            end row in the MFE to print (default=2)
        pprint : boolean
            If True (default), the parameters are printed, else a
            dictionary is returned.

        Returns
        -------
        mfeData : tuple or None
            if ``pprint`` is True, it returns None

        See Also
        --------
        zOptimize() :
            To update the merit function prior to calling ``ipzGetMFE()``,
            call ``zOptimize()`` with the number of cycles set to -1
        """
        if pprint:
            print("Multi-Function Editor contents:")
            print('{:^8}{:>6}{:>6}{:>8}{:>8}{:>8}{:>8}{:>8}{:>10}{:>8}{:>10}'
                  .format("Oper", "int1", "int2", "data1", "data2", "data3", "data4", 
                          "data5", "Target", "Weight", "Value"))
        else:
            mfed = _co.namedtuple('MFEdata', ['Oper', 'int1', 'int2', 'data1', 'data2', 
                                              'data3', 'data4', 'data5', 'Target',
                                              'Weight', 'Value'])
            mfeData = []
        for i in range(start_row, end_row + 1):
            #opr, tgt = self.zGetOperand(i, 1), self.zGetOperand(i, 8)
            #wgt, val = self.zGetOperand(i, 9), self.zGetOperand(i, 10)
            
            odata = self.zGetOperandRow(row=i)
            opr, i1, i2, d1, d2, d3, d4, d5, d6, tgt, wgt, val, per = odata   

            if pprint:
                if isinstance(i1, str):
                    print('{:^8}{}'
                          .format(opr, i1, tgt, wgt, val))
                else:
                    print('{:^8}{:>6.2f}{:>6.2f}{:>8.4f}{:>8.4f}{:>8.4f}{:>8.4f}{:>8.4f}{:>10.6f}{:>8.4f}{:>10.6f}'
                          .format(opr, i1, i2, d1, d2, d3, d4, d5, tgt, wgt, val))
            else:
                data = mfed._make([opr, i1, i2, d1, d2, d3, d4, d5, tgt, wgt, val])
                mfeData.append(data)
        if not pprint:
            return tuple(mfeData)

    def ipzGetPupil(self, pprint=True):
        """Print/ return pupil data in human readable form

        Parameters
        ----------
        pprint : boolean
            If True (default), the parameters are printed, else a
            dictionary is returned

        Returns
        -------
        Print or return dictionary containing pupil information in human
        readable form that meant to be used in interactive environment.
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
        pupil['Entrance pupil position (from surface 1)'] = pupilData[3]
        pupil['Exit pupil diameter'] = pupilData[4]
        pupil['Exit pupil position (from IMA)'] = pupilData[5]
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
        pprint : boolean
            If True (default), the parameters are printed, else a
            dictionary is returned

        Returns
        -------
        Print or return dictionary containing system aperture information
        in human readable form that meant to be used in interactive
        environment.
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

    def ipzGetSurfaceData(self, surfNum, pprint=True):
        """Print or return basic (not all) surface data in human readable
        form

        Parameters
        ----------
        surfNum : integer
            surface number
        pprint : boolean
            If True (default), the parameters are printed, else a
            dictionary is returned.

        Returns
        -------
        Print or return dictionary containing basic surface data (radius
        of curvature, thickness, glass, semi-diameter, and conic) in human
        readable form that meant to be used in interactive environment.
        """
        surfdata = {}
        surfdata['Radius of curvature'] = 1.0/self.zGetSurfaceData(surfNum, 2)
        surfdata['Thickness'] = self.zGetSurfaceData(surfNum, 3)
        surfdata['Glass'] = self.zGetSurfaceData(surfNum, 4)
        surfdata['Semi-diameter'] = self.zGetSurfaceData(surfNum, 5)
        surfdata['Conic'] = self.zGetSurfaceData(surfNum, 6)
        if pprint:
            _print_dict(surfdata)
        else:
            return surfdata

    def ipzGetLDE(self, num=None):
        """Prints the sequential mode LDE data into the IPython cell

        Usage: ``ipzGetLDE()``

        Parameters
        ----------
        num : integer, optional 
            if not None, sufaces upto surface number equal to `num` 
            will be retrieved

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
        numSurf2show = num if num is not None else numSurf 
        line_list = _readLinesFromFile(_openFile(textFileName))
        for line_num, line in enumerate(line_list):
            sectionString = ("SURFACE DATA SUMMARY:") # to use re later
            if line.rstrip()== sectionString:
                for i in range(numSurf2show + 4): # 1 object surf + 3 extra lines before actual data
                    lde_line = line_list[line_num + i].rstrip()
                    print(lde_line)
                break
        else:
            raise Exception("Could not find string '{}' in Prescription file."
            " \n\nPlease check if there is a mismatch in text encoding between"
            " Zemax and PyZDDE, ``Surface Data`` is enabled in prescription"
            " file, and the mode is not pure NSC".format(sectionString))
        _deleteFile(textFileName)


    def ipzGetFieldData(self):
        """Prints formatted field data in IPython QtConsole or Notebook  
        """
        fieldType = {0 : 'Angles in degrees', 
                     1 : 'Object height', 
                     2 : 'Paraxial image height', 
                     3 : 'Real image height'}
        fieldNormalization = {0 : 'Radial', 1 : 'Rectangular'}
        fieldMetaData = self.zGetField(0)
        fieldMeta = {}
        fieldMeta['Type'] = fieldType[fieldMetaData.type]
        fieldMeta['Number of Fields'] = fieldMetaData.numFields
        fieldMeta['Max X'] = fieldMetaData.maxX
        fieldMeta['Max Y'] = fieldMetaData.maxY
        fieldMeta['Field Normalization'] = fieldNormalization[fieldMetaData.normMethod]
        _print_dict(fieldMeta)
        print(("{:^8}{:^8}{:^8}{:^8}{:^8}{:^8}{:^8}{:^8}"
               .format('X', 'Y', 'Weight', 'VDX', 'VDY', 'VCX', 'VCY', 'VAN')))
        for each in self.zGetFieldTuple():
            print(("{:< 8.2f}{:< 8.2f}{:<8.4f}{:<8.4f}{:<8.4f}"
                   "{:<8.4f}{:<8.4f}{:<8.4f}"
                   .format(each.xf, each.yf, each.wgt, each.vdx, each.vdy, 
                           each.vcx, each.vcy, each.van)))

#%% OTHER MODULE HELPER FUNCTIONS THAT DO NOT REQUIRE A RUNNING ZEMAX SESSION

def numAper(aperConeAngle, rIndex=1.0):
    """Returns the Numerical Aperture (NA) for the associated aperture
    cone angle

    Parameters
    ----------
    aperConeAngle : float
        aperture cone angle, in radians
    rIndex : float
        refractive index of the medium

    Returns
    -------
    na : float
        Numerical Aperture
    """
    return rIndex*_math.sin(aperConeAngle)

def numAper2fnum(na, ri=1.0):
    """Convert numerical aperture (NA) to F-number

    Parameters
    ----------
    na : float
        Numerical aperture value
    ri : float
        Refractive index of the medium

    Returns
    -------
    fn : float
        F-number value
    """
    return 1.0/(2.0*_math.tan(_math.asin(na/ri)))

def fnum2numAper(fn, ri=1.0):
    """Convert F-number to numerical aperture (NA)

    Parameters
    ----------
    fn : float
        F-number value
    ri : float
        Refractive index of the medium

    Returns
    -------
    na : float
        Numerical aperture value
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
        to the focal length of the lens for infinite conjugate, or the
        image plane distance, in the same units of length as ``r``
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
       represents the number of annular Fresnel zones in the aperture
       opening [Wolf2011]_, or from the center of the beam to the edge in
       case of a propagating beam [Zemax]_.

    References
    ----------
    .. [Zemax] Zemax manual

    .. [Born&Wolf2011] Principles of Optics, Born and Wolf, 2011
    """
    if approx:
        return (r**2)/(wl*z)
    else:
        return 2.0*(_math.sqrt(z**2 + r**2) - z)/wl

def approx_equal(x, y, tol=macheps):
    """compare two float values using relative difference as measure
    
    Parameters
    ----------
    x, y : floats
        floating point values to be compared
    tol : float
        tolerance (default=`macheps`, which is the difference between 1 and the next 
        representable float. `macheps` is equal to 2^{−23} ≃ 1.19e-07 for 32 bit 
        representation and equal to 2^{−52} ≃ 2.22e-16 for 64 bit representation)
    
    Returns
    -------
    rel_diff : bool
        ``True`` if ``x`` and ``y`` are approximately equal within the tol   
    
    Notes
    -----
    1. relative difference: http://en.wikipedia.org/wiki/Relative_change_and_difference
    3. In future, this function could be replaced by a standard library function. See
       PEP0485 for details. https://www.python.org/dev/peps/pep-0485/
    """
    return abs(x - y) <= max(abs(x), abs(y)) * tol

# scales to SI-meter
#                    mm  , cm  , inch  , m
_zbf_unit_factors = [1e-3, 1e-2, 0.0254, 1]

def zemaxUnitToMeter(zemaxUnitId, value):
    """Converts a zemax unit to SI-meter.

    Parameters
    ----------
    zemaxUnitId : int
        0: mm
        1: cm
        2: inch
        3: m

    Returns
    -------
    value in meter(m)
    """
    return _zbf_unit_factors[zemaxUnitId] * value


def readBeamFile(beamfilename):
    """Read in a Zemax Beam file

    Parameters
    ----------
    beamfilename : string
        the filename of the beam file to read

    Returns
    -------
    version : integer
        the file format version number
    n : 2-tuple, (nx, ny)
        the number of samples in the x and y directions
    ispol : boolean
        is the beam polarized?
    units : integer (0 or 1 or 2 or 3)
        the units of the beam, 0 = mm, 1 = cm, 2 = in, 3 for m
    d : 2-tuple, (dx, dy)
        the x and y grid spacing
    zposition : 2-tuple, (zpositionx, zpositiony)
        the x and y z position of the beam
    rayleigh : 2-tuple, (rayleighx, rayleighy)
        the x and y rayleigh ranges of the beam
    waist : 2-tuple, (waistx, waisty)
        the x and y waists of the beam
    lamda : double
        the wavelength of the beam
    index : double
        the index of refraction in the current medium
    receiver_eff : double
        the receiver efficiency. Zero if fiber coupling is not computed
    system_eff : double
        the system efficiency. Zero if fiber coupling is not computed.
    grid_pos : 2-tuple of lists, (x_matrix, y_matrix)
        lists of x and y positions of the grid defining the beam
    efield : 4-tuple of 2D lists, (Ex_real, Ex_imag, Ey_real, Ey_imag)
        a tuple containing two dimensional lists with the real and
        imaginary parts of the x and y polarizations of the beam

    """
    _warnings.warn('Function readBeamFile() has been moved to zfileutils module. '
                   'Please update code and use the zfileutils module. This function '
                   'will be removed from the zdde module in future.')
    return _zfu.readBeamFile(beamfilename)

def writeBeamFile(beamfilename, version, n, ispol, units, d, zposition, rayleigh,
                 waist, lamda, index, receiver_eff, system_eff, efield):
    """Write a Zemax Beam file

    Parameters
    ----------
    beamfilename : string
        the filename of the beam file to read
    version : integer
        the file format version number
    n : 2-tuple, (nx, ny)
        the number of samples in the x and y directions
    ispol : boolean
        is the beam polarized?
    units : integer
        the units of the beam, 0 = mm, 1 = cm, 2 = in, 3  = m
    d : 2-tuple, (dx, dy)
        the x and y grid spacing
    zposition : 2-tuple, (zpositionx, zpositiony)
        the x and y z position of the beam
    rayleigh : 2-tuple, (rayleighx, rayleighy)
        the x and y rayleigh ranges of the beam
    waist : 2-tuple, (waistx, waisty)
        the x and y waists of the beam
    lamda : double
        the wavelength of the beam
    index : double
        the index of refraction in the current medium
    receiver_eff : double
        the receiver efficiency. Zero if fiber coupling is not computed
    system_eff : double
        the system efficiency. Zero if fiber coupling is not computed.
    efield : 4-tuple of 2D lists, (Ex_real, Ex_imag, Ey_real, Ey_imag)
        a tuple containing two dimensional lists with the real and
        imaginary parts of the x and y polarizations of the beam

    Returns
    -------
    status : integer
        0 = success; -997 = file write failure; -996 = couldn't convert
        data to integer, -995 = unexpected error.
    """
    _warnings.warn('Function writeBeamFile() has been moved to zfileutils module. '
                   'Please update code and use the zfileutils module. This function '
                   'will be removed from the zdde module in future')
    return _zfu.writeBeamFile(beamfilename, version, n, ispol, units, d, zposition, 
                              rayleigh, waist, lamda, index, receiver_eff, system_eff, efield)

def showMessageBox(msg, title='', msgtype='info'):
    """helper function (blocking) to show a simple Tkinter based messagebox. 
    
    Note that the call is a blocking call, halting the execution of the 
    program till an action (click of the OK button) is performed by the
    user.
    
    Parameters
    ----------
    msg : string
        text to be displayed as a message (can occupy multiple lines)
    title : string, optional
        the text to be displayed in the title bar of a message box
    msgtype : string, optional
        'info', 'warn', or 'error' to indicate the type of message.
        If no `msgtype` not given, or the string is not one of the
        above, an info type messagebox is displayed.
        
    Returns
    -------
    None
    """
    _tk.Tk().withdraw()
    msg_func_dict = { 'info': _MessageBox.showinfo,
                      'warn': _MessageBox.showwarning,
                     'error': _MessageBox.showerror}
    try:
        func = msg_func_dict[msgtype]
    except KeyError:
        func = msg_func_dict['info']
    func(title=title, message=msg)

#%% Helper functions to process data from ZEMAX DDE server. 
# This is especially convenient for processing replies from Zemax for 
# those function calls that a known data structure. These functions are 
# mainly used intenally and may not be exposed directly.

def _regressLiteralType(x):
    """The function returns the literal with its proper type, such as int,
    float, or string from the input string x

    Examples
    --------
    >>> _regressLiteralType("1")->1
    >>> _regressLiteralType("1.0")->1.0
    >>> _regressLiteralType("1e-3")->0.001
    >>> _regressLiteralType("YUV")->'YUV'
    >>> _regressLiteralType("YO8")->'YO8'
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

    Parameters
    ----------
    fileName : string
        file name of file to be deleted with full path
    n : integer
        number of times to attempt before giving up

    Returns
    -------
    status : bool
        True  = file deleting successful;
        False = reached maximum number of attempts, without deleting file.

    Notes
    -----
    It assumes that the file with filename actually exist and doesn't do
    any error checking on its existance. This is OK as this function is
    for internal use only.
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

def _deleteFilesCreatedDuringSession(self):
    """Helper function to clean up files creatd by PyZDDE during a session.
    Examples of such files include configuration files, etc.
    """
    filesToDelete = self._filesCreated
    filesNotDeleted = set()
    for filename in filesToDelete:
        if not _deleteFile(filename):
            filesNotDeleted.add(filename)
    remaining = filesToDelete.intersection(filesNotDeleted)
    self._filesCreated = remaining

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
    elif column in (2,3): # if thre is a comment, it will be in column 2
        #return int(float(rs))
        return _regressLiteralType(rs)
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
    if code in  (102,103, 104,105,106,107,108,109,110,202,203): # unexpected (identified) cases
        sysPropData = reply
    elif code in (16,17,23,40,41,42,43): # string
        sysPropData = reply.rstrip()    #str(reply)
    elif code in (11,13,24,53,54,55,56,60,61,62,63,71,72,73,77,78): # floats
        sysPropData = float(reply)
    else:
        sysPropData = int(float(reply))      # integer
    return sysPropData

def _process_get_set_Tol(operNum,reply):
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

    Notes
    -----
    This is just a wrapper around the open function.
    It is the responsibility of the calling function to close the file by
    calling the ``close()`` method of the file object. Alternatively use
    either use a with/as context to close automatically or use
    ``_readLinesFromFile()`` or ``_getDecodedLineFromFile()`` that uses a
    with context manager to handle exceptions and file close.
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

    This function emulates the functionality of ``readlines()`` method of file
    objects. The caller doesn't have to explicitly close the file as it is
    handled in ``_getDecodedLineFromFile()`` function.

    Parameters
    ----------
    fileObj : file object returned by ``open()`` method

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

def _get2DList(line_list, start_line, number_of_lines, 
               startCol=None, endCol=None, stride=None):
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
    startCol : integer, optional 
        the column number to start reading in each row (similar to list 
        slicing pattern). Default is `None`
    endCol : integer, optional
        the end column number upto which (but excluding `endCol`) to read
        in each row (similar to list slicing pattern). Default is `None`
    stride : integer, optional  
        stride along each column (similar to list slicing pattern). 
        Default is `None`

    Returns
    -------
    data : list
        data is a 2-d list of float type data read from the lines in
        line_list
    """
    data = []
    end_line = start_line + number_of_lines - 1
    for lineNum, row in enumerate(line_list):
        if start_line <= lineNum <= end_line:
            data.append([float(i) for i in row.split()][startCol:endCol:stride])
    return data

def _transpose2Dlist(mat):
    """transpose a matrix that is constructed as a list of lists in pure
    Python

    The inner lists represents rows.

    Parameters
    ----------
    mat : list of lists (2-d list)
        the 2D list represented as
        | [[a_00, a_01, ..., a_0n],
        |  [a_10, a_11, ..., a_1n],
        |            ...          ,
        |  [a_m0, a_m1, ..., a_mn]]

    Returns
    -------
    matT : list of lists (2-d list)
        transposed of ``mat``

    Notes
    -----
    The function assumes that all the inner lists are of the same lengths.
    It doesn't do any error checking.
    """
    cols = len(mat[0])
    matT = []
    for i in range(cols):
        matT.append([row[i] for row in mat])
    return matT

def _getRePatPosInLineList(line_list, re_pattern):
    """internal helper function for retrieving the positions of specific
    patterns in a list of lines read from a file

    Parameters
    ----------
    line_list : list
        list of lines read from a file
    re_pattern : string
        regular expression pattern to loop for

    Returns
    -------
    positions : list
        the list containing the position of the pattern in the input list
    """
    positions = []
    for line_number, line in enumerate(line_list):
        if _re.search(re_pattern, line):
            positions.append(line_number)
    return positions

def _txtAndSettingsToUse(self, txtFile, settingsFile, anaType):
    """internal helper function for use by zGet type of functions
    that call ``zGetTextFile()``, to decide the type of settings
    file and settings flag to use

    Parameters
    ----------
    self : object
        pyzdde link object
    txtFile : string
        text file that may have been passed by the user
    settingsFile : string
        settings file that may have been passed by the user
    anaType : string
        3-letter analysis code

    Returns
    -------
    cfgFile : string
        full name and path of the configuration/ settings file to
        use for calling ``zGetTextFile()``
    getTextFlag : integer
        flag to be used for calling ``zGetTextFile()``
    """
    # note to the developer -- maintain exactly same keys in both
    # txtFileDict and anaCfgDict. Note that some analysis have common
    # txt file and settings files associated with them. Of course they
    # may be changed if required in future.
    txtFileDict =  {'Pop':'popData.txt',
                    'Hcs':'huygensPsfCSAnalysisFile.txt', # Huygens PSF cross-section
                    'Hps':'huygensPsfAnalysisFile.txt', # Huygens PSF
                    'Hmf':'huygensMtfAnalysisFile.txt', # Huygens MTF
                    'Pcs':'fftPsfCSAnalysisFile.txt', # FFT PSF cross-section
                    'Fps':'fftPsfAnalysisFile.txt', # FFT PSF
                    'Mtf':'fftMtfAnalysisFile.txt', # FFT MTF
                    'Sei':'seidelAberrationFile.txt', # Seidel aberration coefficients
                    'Pre':'prescriptionFile.txt',   # Prescription
                    'Sim':'imageSimulationAnalysisFile.txt', # Image Simulation
                    'Zfr':'zernikeFringeAnalysisFile.txt',   # Zernike Fringe coefficients
                    'Zst':'zernikeStandardAnalysisFile.txt', # Zernike Standard coefficients
                    'Zat':'zernikeAnnularAnalysisFile.txt',  # Zernike Annular coefficients
                    'Dvw':'detectorViewerFile.txt',  # NSC detector viewer         
                    }

    anaCfgDict  = {'Pop':'_pyzdde_POP.CFG',
                   'Hcs':'_pyzdde_HUYGENSPSFCS.CFG',
                   'Hps':'_pyzdde_HUYGENSPSF.CFG',
                   'Hmf':'_pyzdde_HUYGENSMTF.CFG',
                   'Pcs':'_pyzdde_FFTPSFCS.CFG',
                   'Fps':'_pyzdde_FFTPSF.CFG',
                   'Mtf':'_pyzdde_FFTMTF.CFG',
                   'Sei':'None',
                   'Pre':'None',  # change this to the appropriate file when implemented
                   'Sim':'_pyzdde_IMGSIM.CFG',
                   'Zfr':'_pyzdde_ZFR.CFG',  # Note that currently MODIFYSETTINGS
                   'Zst':'_pyzdde_ZST.CFG',  # is not supported of Aberration
                   'Zat':'_pyzdde_ZAT.CFG',  # coefficients by Zemax extensions
                   'Dvw':'_pyzdde_DVW.CFG',  # NSC detector viewer
                   }
    assert txtFileDict.keys() == anaCfgDict.keys(), \
           "Dicts don't have matching keys" # for code integrity
    assert anaType in anaCfgDict

    #fdir = _os.path.dirname(_os.path.realpath(__file__))
    fdir = _os.path.dirname(self.zGetFile())
    if txtFile != None:
        textFileName = txtFile
    else:
        textFileName = _os.path.join(fdir, txtFileDict[anaType])
    if settingsFile:
        cfgFile = settingsFile
        getTextFlag = 1
    else:
        f = _os.path.splitext(self.zGetFile())[0] + anaCfgDict[anaType]
        if _checkFileExist(f): # use "*_pyzdde_XXX.CFG" settings file
            cfgFile = f
            getTextFlag = 1
        else: # use default settings file
            cfgFile = ''
            getTextFlag = 0
    return textFileName, cfgFile, getTextFlag

#
#
if __name__ == "__main__":
    print("Please import this module as 'import pyzdde.zdde as pyz' ")
    _sys.exit(0)
