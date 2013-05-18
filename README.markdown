PyZDDE is a Python toolbox for communicating with ZEMAX using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. It is similar to the MZDDE toolbox, in Matlab, developed by Derek Griffith at CSIR (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx).



##PYTHON VERSION and LIBRARIES USED:

1.   Python 2.7.3, 32 bit version (required)

2.   PyWin32, build 214 (or earlier), 32 bit version (required) [See bullet 2 of the "CURRENT ISSUES" section for issues] 

3.   Matplotlib (optional, used in some of the example programs)


##CURENT ISSUES:

1.   The software is CURRENTLY NOT COMPLETE! Quite a number of functions are not yet implemented.

2.   Due to known a bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. The known issue has been discussed and documented at http://sourceforge.net/mailarchive/message.php?msg_id=28828321


##FILES AND DIRECTORIES:

PyZDDE                     : Top-level folder for the toolbox

PyZDDE/LICENSE      : License file (MIT License)

PyZDDE/pyZDDE.py   : Main file containing the pyZDDE code

PyZDDE/Test             : Folder containing the unit-test code for pyZDDE 

pyZDDE/Examples     : Folder containing example codes that uses pyZDDE

pyZDDE/ZMXFILES    : Folder containing a few Zemax Lens files for testing the functionality of the toolbox.


