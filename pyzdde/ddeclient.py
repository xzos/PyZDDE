# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        ddeclient.py
# Purpose:     DDE Management Library (DDEML) client application for communicating
#              with Zemax
#
# Notes:       This code has been adapted from David Naylor's dde-client code from
#              ActiveState's Python recipes (Revision 1).
# Author of original Code:   David Naylor, Apr 2011
# Modified by Indranil Sinharoy
# Copyright:   (c) David Naylor
# Licence:     New BSD license (Please see the file Notice.txt for further details)
# Website:     http://code.activestate.com/recipes/577654-dde-client/
#-------------------------------------------------------------------------------
from __future__ import print_function
import sys
from ctypes import c_int, c_double, c_char_p, c_void_p, c_ulong, c_char, pointer, cast
from ctypes import windll, byref, create_string_buffer, Structure, sizeof
from ctypes import POINTER, WINFUNCTYPE
from ctypes.wintypes import BOOL, HWND, MSG, DWORD, BYTE, INT, LPCWSTR, UINT, ULONG, LPCSTR

# DECLARE_HANDLE(name) typedef void *name;
HCONV     = c_void_p  # = DECLARE_HANDLE(HCONV)
HDDEDATA  = c_void_p  # = DECLARE_HANDLE(HDDEDATA)
HSZ       = c_void_p  # = DECLARE_HANDLE(HSZ)
LPBYTE    = c_char_p  # POINTER(BYTE)
LPDWORD   = POINTER(DWORD)
LPSTR     = c_char_p
ULONG_PTR = c_ulong

# See windows/ddeml.h for declaration of struct CONVCONTEXT
PCONVCONTEXT = c_void_p

# DDEML errors
DMLERR_NO_ERROR            = 0x0000  # No error
DMLERR_ADVACKTIMEOUT       = 0x4000  # request for synchronous advise transaction timed out
DMLERR_DATAACKTIMEOUT      = 0x4002  # request for synchronous data transaction timed out
DMLERR_DLL_NOT_INITIALIZED = 0x4003  # DDEML functions called without iniatializing
DMLERR_EXECACKTIMEOUT      = 0x4006  # request for synchronous execute transaction timed out
DMLERR_NO_CONV_ESTABLISHED = 0x400a  # client's attempt to establish a conversation has failed (can happen during DdeConnect)
DMLERR_POKEACKTIMEOUT      = 0x400b  # A request for a synchronous poke transaction has timed out.
DMLERR_POSTMSG_FAILED      = 0x400c  # An internal call to the PostMessage function has failed.
DMLERR_SERVER_DIED         = 0x400e

# Predefined Clipboard Formats
CF_TEXT         =  1
CF_BITMAP       =  2
CF_METAFILEPICT =  3
CF_SYLK         =  4
CF_DIF          =  5
CF_TIFF         =  6
CF_OEMTEXT      =  7
CF_DIB          =  8
CF_PALETTE      =  9
CF_PENDATA      = 10
CF_RIFF         = 11
CF_WAVE         = 12
CF_UNICODETEXT  = 13
CF_ENHMETAFILE  = 14
CF_HDROP        = 15
CF_LOCALE       = 16
CF_DIBV5        = 17
CF_MAX          = 18

# DDE constants for wStatus field
DDE_FACK          = 0x8000
DDE_FBUSY         = 0x4000
DDE_FDEFERUPD     = 0x4000
DDE_FACKREQ       = 0x8000
DDE_FRELEASE      = 0x2000
DDE_FREQUESTED    = 0x1000
DDE_FAPPSTATUS    = 0x00FF
DDE_FNOTPROCESSED = 0x0000

DDE_FACKRESERVED  = (~(DDE_FACK | DDE_FBUSY | DDE_FAPPSTATUS))
DDE_FADVRESERVED  = (~(DDE_FACKREQ | DDE_FDEFERUPD))
DDE_FDATRESERVED  = (~(DDE_FACKREQ | DDE_FRELEASE | DDE_FREQUESTED))
DDE_FPOKRESERVED  = (~(DDE_FRELEASE))

# DDEML Transaction class flags
XTYPF_NOBLOCK        = 0x0002
XTYPF_NODATA         = 0x0004
XTYPF_ACKREQ         = 0x0008

XCLASS_MASK          = 0xFC00
XCLASS_BOOL          = 0x1000
XCLASS_DATA          = 0x2000
XCLASS_FLAGS         = 0x4000
XCLASS_NOTIFICATION  = 0x8000

XTYP_ERROR           = (0x0000 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK)
XTYP_ADVDATA         = (0x0010 | XCLASS_FLAGS)
XTYP_ADVREQ          = (0x0020 | XCLASS_DATA | XTYPF_NOBLOCK)
XTYP_ADVSTART        = (0x0030 | XCLASS_BOOL)
XTYP_ADVSTOP         = (0x0040 | XCLASS_NOTIFICATION)
XTYP_EXECUTE         = (0x0050 | XCLASS_FLAGS)
XTYP_CONNECT         = (0x0060 | XCLASS_BOOL | XTYPF_NOBLOCK)
XTYP_CONNECT_CONFIRM = (0x0070 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK)
XTYP_XACT_COMPLETE   = (0x0080 | XCLASS_NOTIFICATION )
XTYP_POKE            = (0x0090 | XCLASS_FLAGS)
XTYP_REGISTER        = (0x00A0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK )
XTYP_REQUEST         = (0x00B0 | XCLASS_DATA )
XTYP_DISCONNECT      = (0x00C0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK )
XTYP_UNREGISTER      = (0x00D0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK )
XTYP_WILDCONNECT     = (0x00E0 | XCLASS_DATA | XTYPF_NOBLOCK)
XTYP_MONITOR         = (0x00F0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK)

XTYP_MASK            = 0x00F0
XTYP_SHIFT           = 4

# DDE Timeout constants
TIMEOUT_ASYNC        = 0xFFFFFFFF

# DDE Application command flags / Initialization flag (afCmd)
APPCMD_CLIENTONLY    = 0x00000010

# Code page for rendering string.
CP_WINANSI      = 1004    # default codepage for windows & old DDE convs.
CP_WINUNICODE   = 1200

# Declaration
DDECALLBACK = WINFUNCTYPE(HDDEDATA, UINT, UINT, HCONV, HSZ, HSZ, HDDEDATA,  ULONG_PTR, ULONG_PTR)

# PyZDDE specific globals
number_of_apps_communicating = 0  # to keep an account of the number of zemax
                                  # server objects --'ZEMAX', 'ZEMAX1' etc

class CreateServer(object):
    """This is really just an interface class so that PyZDDE can use either the
    current dde code or the pywin32 transparently. This object is created only
    once. The class name cannot be anything else if compatibility has to be maintained
    between pywin32 and this dde code.
    """
    def __init__(self):
        self.serverName = 'None'

    def Create(self, client):
        """Set a DDE client that will communicate with the DDE server

        Parameters
        ----------
        client : string
            Name of the DDE client, most likely this will be 'ZCLIENT'
        """
        self.clientName = client  # shall be used in `CreateConversation`

    def Shutdown(self, createConvObj):
        """The shutdown should ideally be requested only once per CreateConversation
        object by the PyZDDE module, but for ALL CreateConversation objects, if there
        are more than one. If multiple CreateConversation objects were created and
        then not cleared, there will be memory leak, and eventually the program will
        error out when run multiple times

        Parameters
        ----------
        createConvObj : CreateConversation object

        Exceptions
        ----------
        An exception occurs if a Shutdown is attempted with a CreateConvObj that
        doesn't have a conversation object (connection with ZEMAX established)
        """
        global number_of_apps_communicating
        #print("Shutdown requested by {}".format(repr(createConvObj))) # for debugging
        if number_of_apps_communicating > 0:
            #print("Deleting object ...") # for debugging
            createConvObj.ddec.__del__()
            number_of_apps_communicating -=1


class CreateConversation(object):
    """This is really just an interface class so that PyZDDE can use either the
    current dde code or the pywin32 transparently.

    Multiple objects of this type may be instantiated depending upon the
    number of simultaneous channels of communication with Zemax that the user
    program wants to establish using `ln = pyz.PyZDDE()` followed by `ln.zDDEInit()`
    calls.
    """
    def __init__(self, ddeServer):
        """
        Parameters
        ----------
        ddeServer :
            d
        """
        self.ddeClientName = ddeServer.clientName
        self.ddeServerName = 'None'
        self.ddetimeout = 50    # default dde timeout = 50 seconds

    def ConnectTo(self, appName, data=None):
        """Exceptional error is handled in zdde Init() method, so the exception
        must be re-raised"""
        global number_of_apps_communicating
        self.ddeServerName = appName
        try:
            self.ddec = DDEClient(self.ddeServerName, self.ddeClientName) # establish conversation
        except DDEError:
            raise
        else:
            number_of_apps_communicating +=1
        #print("Number of apps communicating: ", number_of_apps_communicating) # for debugging

    def Request(self, item, timeout=None):
        """Request DDE client
        timeout in seconds
        Note ... handle the exception within this function.
        """
        if not timeout:
            timeout = self.ddetimeout
        try:
            reply = self.ddec.request(item, int(timeout*1000)) # convert timeout into milliseconds
        except DDEError:
            err_str = str(sys.exc_info()[1])
            error = err_str[err_str.find('err=')+4:err_str.find('err=')+10]
            if error == hex(DMLERR_DATAACKTIMEOUT):
                print("TIMEOUT REACHED. Please use a higher timeout.\n")
            if (sys.version_info > (3, 0)): #this is only evaluated in case of an error
                reply = b'-998' #Timeout error value
            else:
                reply = '-998' #Timeout error value
        return reply

    def RequestArrayTrace(self, ddeRayData, timeout=None):
        """Request bulk ray tracing

        Parameters
        ----------
        ddeRayData : the ray data for array trace
        """
        pass
        # TO DO!!!
        # 1. Assign proper timeout as in Request() function
        # 2. Create the rayData structure conforming to ctypes structure
        # 3. Process the reply and return ray trace data
        # 4. Handle errors
        #reply = self.ddec.poke("RayArrayData", rayData, timeout)

    def SetDDETimeout(self, timeout):
        """Set DDE timeout
        timeout : timeout in seconds
        """
        self.ddetimeout = timeout

    def GetDDETimeout(self):
        """Returns the current timeout value in seconds
        """
        return self.ddetimeout


def get_winfunc(libname, funcname, restype=None, argtypes=(), _libcache={}):
    """Retrieve a function from a library/DLL, and set the data types."""
    if libname not in _libcache:
        _libcache[libname] = windll.LoadLibrary(libname)
    func = getattr(_libcache[libname], funcname)
    func.argtypes = argtypes
    func.restype = restype
    return func

class DDE(object):
    """Object containing all the DDEML functions"""
    AccessData         = get_winfunc("user32", "DdeAccessData",          LPBYTE,   (HDDEDATA, LPDWORD))
    ClientTransaction  = get_winfunc("user32", "DdeClientTransaction",   HDDEDATA, (LPBYTE, DWORD, HCONV, HSZ, UINT, UINT, DWORD, LPDWORD))
    Connect            = get_winfunc("user32", "DdeConnect",             HCONV,    (DWORD, HSZ, HSZ, PCONVCONTEXT))
    CreateDataHandle   = get_winfunc("user32", "DdeCreateDataHandle",    HDDEDATA, (DWORD, LPBYTE, DWORD, DWORD, HSZ, UINT, UINT))
    CreateStringHandle = get_winfunc("user32", "DdeCreateStringHandleW", HSZ,      (DWORD, LPCWSTR, UINT))  # Unicode version
    #CreateStringHandle = get_winfunc("user32", "DdeCreateStringHandleA", HSZ,      (DWORD, LPCSTR, UINT))  # ANSI version
    Disconnect         = get_winfunc("user32", "DdeDisconnect",          BOOL,     (HCONV,))
    GetLastError       = get_winfunc("user32", "DdeGetLastError",        UINT,     (DWORD,))
    Initialize         = get_winfunc("user32", "DdeInitializeW",         UINT,     (LPDWORD, DDECALLBACK, DWORD, DWORD)) # Unicode version of DDE initialize
    #Initialize         = get_winfunc("user32", "DdeInitializeA",         UINT,     (LPDWORD, DDECALLBACK, DWORD, DWORD)) # ANSI version of DDE initialize
    FreeDataHandle     = get_winfunc("user32", "DdeFreeDataHandle",      BOOL,     (HDDEDATA,))
    FreeStringHandle   = get_winfunc("user32", "DdeFreeStringHandle",    BOOL,     (DWORD, HSZ))
    QueryString        = get_winfunc("user32", "DdeQueryStringA",        DWORD,    (DWORD, HSZ, LPSTR, DWORD, c_int)) # ANSI version of QueryString
    UnaccessData       = get_winfunc("user32", "DdeUnaccessData",        BOOL,     (HDDEDATA,))
    Uninitialize       = get_winfunc("user32", "DdeUninitialize",        BOOL,     (DWORD,))

class DDEError(RuntimeError):
    """Exception raise when a DDE error occures."""
    def __init__(self, msg, idInst=None):
        if idInst is None:
            RuntimeError.__init__(self, msg)
        else:
            RuntimeError.__init__(self, "%s (err=%s)" % (msg, hex(DDE.GetLastError(idInst))))

class DDEClient(object):
    """The DDEClient class.

    Use this class to create and manage a connection to a service/topic.  To get
    classbacks subclass DDEClient and overwrite callback."""

    def __init__(self, service, topic):
        """Create a connection to a service/topic."""
        self._idInst = DWORD(0) # application instance identifier.
        self._hConv = HCONV()

        self._callback = DDECALLBACK(self._callback)
        # Initialize and register application with DDEML
        res = DDE.Initialize(byref(self._idInst), self._callback, APPCMD_CLIENTONLY, 0)
        if res != DMLERR_NO_ERROR:
            raise DDEError("Unable to register with DDEML (err=%s)" % hex(res))

        hszServName = DDE.CreateStringHandle(self._idInst, service, CP_WINUNICODE)
        hszTopic = DDE.CreateStringHandle(self._idInst, topic, CP_WINUNICODE)
        # Try to establish conversation with the Zemax server
        self._hConv = DDE.Connect(self._idInst, hszServName, hszTopic, PCONVCONTEXT())
        DDE.FreeStringHandle(self._idInst, hszTopic)
        DDE.FreeStringHandle(self._idInst, hszServName)
        if not self._hConv:
            raise DDEError("Unable to establish a conversation with server", self._idInst)

    def __del__(self):
        """Cleanup any active connections and free all DDEML resources."""
        if self._hConv:
            DDE.Disconnect(self._hConv)
        if self._idInst:
            DDE.Uninitialize(self._idInst)

    def advise(self, item, stop=False):
        """Request updates when DDE data changes."""
        hszItem = DDE.CreateStringHandle(self._idInst, item, CP_WINUNICODE)
        hDdeData = DDE.ClientTransaction(LPBYTE(), 0, self._hConv, hszItem, CF_TEXT, XTYP_ADVSTOP if stop else XTYP_ADVSTART, TIMEOUT_ASYNC, LPDWORD())
        DDE.FreeStringHandle(self._idInst, hszItem)
        if not hDdeData:
            raise DDEError("Unable to %s advise" % ("stop" if stop else "start"), self._idInst)
        DDE.FreeDataHandle(hDdeData)

    def execute(self, command):
        """Execute a DDE command."""
        pData = c_char_p(command)
        cbData = DWORD(len(command) + 1)
        hDdeData = DDE.ClientTransaction(pData, cbData, self._hConv, HSZ(), CF_TEXT, XTYP_EXECUTE, TIMEOUT_ASYNC, LPDWORD())
        if not hDdeData:
            raise DDEError("Unable to send command", self._idInst)
        DDE.FreeDataHandle(hDdeData)

    def request(self, item, timeout=5000):
        """Request data from DDE service."""
        hszItem = DDE.CreateStringHandle(self._idInst, item, CP_WINUNICODE)
        #hDdeData = DDE.ClientTransaction(LPBYTE(), 0, self._hConv, hszItem, CF_TEXT, XTYP_REQUEST, timeout, LPDWORD())
        pdwResult = DWORD(0)
        hDdeData = DDE.ClientTransaction(LPBYTE(), 0, self._hConv, hszItem, CF_TEXT, XTYP_REQUEST, timeout, byref(pdwResult))
        DDE.FreeStringHandle(self._idInst, hszItem)
        if not hDdeData:
            raise DDEError("Unable to request item", self._idInst)

        if timeout != TIMEOUT_ASYNC:
            pdwSize = DWORD(0)
            pData = DDE.AccessData(hDdeData, byref(pdwSize))
            if not pData:
                DDE.FreeDataHandle(hDdeData)
                raise DDEError("Unable to access data in request function", self._idInst)
            DDE.UnaccessData(hDdeData)
        else:
            pData = None
        DDE.FreeDataHandle(hDdeData)
        return pData

    def poke(self, item, data, timeout=5000):
        """Poke (unsolicited) data to DDE server"""
        hszItem = DDE.CreateStringHandle(self._idInst, item, CP_WINUNICODE)
        pData = c_char_p(data)
        cbData = DWORD(len(data) + 1)
        pdwResult = DWORD(0)
        #hData = DDE.CreateDataHandle(self._idInst, data, cbData, 0, hszItem, CP_WINUNICODE, 0)
        #hDdeData = DDE.ClientTransaction(hData, -1, self._hConv, hszItem, CF_TEXT, XTYP_POKE, timeout, LPDWORD())
        hDdeData = DDE.ClientTransaction(pData, cbData, self._hConv, hszItem, CF_TEXT, XTYP_POKE, timeout, byref(pdwResult))
        DDE.FreeStringHandle(self._idInst, hszItem)
        #DDE.FreeDataHandle(dData)
        if not hDdeData:
            print("Value of pdwResult: ", pdwResult)
            raise DDEError("Unable to poke to server", self._idInst)

        if timeout != TIMEOUT_ASYNC:
            pdwSize = DWORD(0)
            pData = DDE.AccessData(hDdeData, byref(pdwSize))
            if not pData:
                DDE.FreeDataHandle(hDdeData)
                raise DDEError("Unable to access data in poke function", self._idInst)
            # TODO: use pdwSize
            DDE.UnaccessData(hDdeData)
        else:
            pData = None
        DDE.FreeDataHandle(hDdeData)
        return pData

    def callback(self, value, item=None):
        """Callback function for advice."""
        print("callback: %s: %s" % (item, value))

    def _callback(self, wType, uFmt, hConv, hsz1, hsz2, hDdeData, dwData1, dwData2):
        """DdeCallback callback function for processing Dynamic Data Exchange (DDE)
        transactions sent by DDEML in response to DDE events

        Parameters
        ----------
        wType    : transaction type (UINT)
        uFmt     : clipboard data format (UINT)
        hConv    : handle to conversation (HCONV)
        hsz1     : handle to string (HSZ)
        hsz2     : handle to string (HSZ)
        hDDedata : handle to global memory object (HDDEDATA)
        dwData1  : transaction-specific data (DWORD)
        dwData2  : transaction-specific data (DWORD)

        Returns
        -------
        ret      : specific to the type of transaction (HDDEDATA)
        """
        if wType == XTYP_ADVDATA:  # value of the data item has changed [hsz1 = topic; hsz2 = item; hDdeData = data]
            dwSize = DWORD(0)
            pData = DDE.AccessData(hDdeData, byref(dwSize))
            if pData:
                item = create_string_buffer('\000' * 128)
                DDE.QueryString(self._idInst, hsz2, item, 128, CP_WINANSI)
                self.callback(pData, item.value)
                DDE.UnaccessData(hDdeData)
                return DDE_FACK
            else:
                print("Error: AccessData returned NULL! (err = %s)"% (hex(DDE.GetLastError(self._idInst))))
        if wType == XTYP_DISCONNECT:
            print("Disconnect notification received from server")

        return 0

def WinMSGLoop():
    """Run the main windows message loop."""
    LPMSG = POINTER(MSG)
    LRESULT = c_ulong
    GetMessage = get_winfunc("user32", "GetMessageW", BOOL, (LPMSG, HWND, UINT, UINT))
    TranslateMessage = get_winfunc("user32", "TranslateMessage", BOOL, (LPMSG,))
    # restype = LRESULT
    DispatchMessage = get_winfunc("user32", "DispatchMessageW", LRESULT, (LPMSG,))

    msg = MSG()
    lpmsg = byref(msg)
    while GetMessage(lpmsg, HWND(), 0, 0) > 0:
        TranslateMessage(lpmsg)
        DispatchMessage(lpmsg)

if __name__ == "__main__":
    pass

