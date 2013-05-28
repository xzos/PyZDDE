PyZDDE is a Python toolbox for communicating with ZEMAX using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. It is similar to the MZDDE toolbox, in Matlab, developed by Derek Griffith at CSIR (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx).



##PYTHON VERSION and LIBRARIES USED:

1.   Python 2.7.3, 32 bit version (required)

2.   PyWin32, build 214 (or earlier), 32 bit version (required) [See bullet 2 of the "CURRENT ISSUES" section for issues] 

3.   Matplotlib (optional, used in some of the example programs)


##CURENT ISSUES:

1.   All the functions for accessing the items that ZEMAX has made available have been written. However, the code base is in rapid development phase. Hence, a distribution version is not available at this point in time. Please download the code to a directory in your computer and add that directory to python search path. SEE  https://www.evernote.com/shard/s82/sh/7ad34692-b2a7-467f-a0be-79116bbbd3cf/b3adc4fa0bb11756b6fdabeade04c49f?noteKey=b3adc4fa0bb11756b6fdabeade04c49f&noteGuid=7ad34692-b2a7-467f-a0be-79116bbbd3cf

2.   Due to known a bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. The known issue has been discussed and documented at http://sourceforge.net/mailarchive/message.php?msg_id=28828321


##FILES AND DIRECTORIES:

PyZDDE                     : Top-level folder for the toolbox 

PyZDDE/LICENSE      : License file (MIT License)

PyZDDE/pyzdde.py   : Main file containing the PyZDDE code. The user must import this module to use the toolbox as "import pyzdde"

PyZDDE/zemaxoperands.py     : File containing the ZEMAX operands class and associated functions

PyZDDE/zemaxbuttons.py       : File containing the ZEMAX buttons class and associated functions.

PyZDDE/Test             : Folder containing the unit-test code for PyZDDE 

PyZDDE/Examples     : Folder containing example codes that uses PyZDDE

PyZDDE/ZMXFILES    : Folder containing a few Zemax Lens files for testing the functionality of the toolbox.


