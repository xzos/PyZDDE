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
import traceback
import pyzdde.zdde as pyz

# Create a DDE link object
ln = pyz.createLink()
try:
    # Write your code to interact with Zemax, for example
    zemaxVer = ln.zGetVersion()
    print("Zemax version: ", zemaxVer)



except Exception, err:
    traceback.print_exc()
finally:
    #Close DDE link
    ln.zDDEClose()
