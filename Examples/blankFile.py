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

#****************** Add PyZDDE to Python search path **************************

PyZDDEPath = 'C:\PyZDDE'  # Assuming PyZDDE was unzipped here!

if PyZDDEPath not in sys.path:
    sys.path.append(PyZDDEPath)
#******************************************************************************

import pyzdde
import zemaxoperands as zo
import zemaxbuttons  as zb

#Create a PyZDDE object
link0 = pyzdde.PyZDDE()

#Initiate the DDE link
status = link0.zDDEInit()











#Close DDE link
link0.zDDEClose()
