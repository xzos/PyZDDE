#-------------------------------------------------------------------------------
# Name:        module_name.py
# Purpose:     A blank file. Use this skeleton file as model for scripts
#              that explicitly adds PyZDDE path to the Python search path.
#
# Created:
# Copyright:
# Licence:     <your licence>
#
# Note:        The coding style shown below is "safe" (using exception handling).
#              In 95% of cases you don't need to use exception handling. PyZDDE
#              is pretty stable.  
#-------------------------------------------------------------------------------
from __future__ import print_function
import sys
import traceback

#****************** Add PyZDDE to Python search path **************************

PyZDDEPath = 'C:\PyZDDE'  # Assuming PyZDDE unzipped here; if not, change the path 
                          # appropriately

if PyZDDEPath not in sys.path:
    sys.path.append(PyZDDEPath)
#******************************************************************************

import pyzdde.zdde as pyz

# Create a DDE link object
link = pyz.createLink()
try:
    # Write your code to interact with Zemax, for example
    zemaxVer = link.zGetVersion()
    print("Zemax version: ", zemaxVer)



except Exception, err:
    traceback.print_exc()
finally:
    #Close DDE link
    link.zDDEClose()
