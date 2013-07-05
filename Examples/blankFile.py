#-------------------------------------------------------------------------------
# Name:        module_name.py
# Purpose:     A blank file
#
# Author:
#
# Created:
# Copyright:
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import sys
import traceback
#****************** Add PyZDDE to Python search path **************************

PyZDDEPath = 'C:\PyZDDE'  # Assuming PyZDDE was unzipped here!

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

    # Write your code to interact with Zemax



except Exception, err:
    traceback.print_exc()
finally:
    #Close DDE link
    link0.zDDEClose()
