#-------------------------------------------------------------------------------
# Name:        module_name.py
# Purpose:     A blank file. Use this skeleton file as model for scripts
#              that explicitly adds PyZDDE path to the Python search path.
#
# Created:
# Copyright:
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from __future__ import print_function
import sys
import traceback

#****************** Add PyZDDE to Python search path **************************

PyZDDEPath = 'C:\PyZDDE'  # Assuming PyZDDE was unzipped here; if not, change the path appropriately

if PyZDDEPath not in sys.path:
    sys.path.append(PyZDDEPath)
#******************************************************************************

import pyzdde.zdde as pyz
import pyzdde.zcodes.zemaxoperands as zo # if required. (use pyz.zo to access module functions)
import pyzdde.zcodes.zemaxbuttons  as zb # if required. (use pyz.zb to access module functions)

# Create a PyZDDE object
link0 = pyz.PyZDDE()
try:
    # Initiate the DDE link
    status = link0.zDDEInit()   # if status == -1, then zDDEInit failed!

    # Write your code to interact with Zemax, for example
    zemaxVer = link0.zGetVersion()
    print("Zemax version: ", zemaxVer)



except Exception, err:
    traceback.print_exc()
finally:
    #Close DDE link
    link0.zDDEClose()
