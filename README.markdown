PyZDDE is a toolbox for communicating with [ZEMAX] (http://www.radiantzemax.com/)  using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. ZEMAX is a leading software tool for design and analysis of optical systems. This toolbox, which implements all of the data items listed in the ZEMAX Extensions chapter of the ZEMAX manual, provides access and control to ZEMAX from Python. It is similar to and very much inspired by the [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) in Matlab which is developed by Derek Griffith at CSIR. However, at this point, it is not as extensive as MZDDE.

### <font color="Brown">Help:</font>
Please refer to the [Wiki] (https://github.com/indranilsinharoy/PyZDDE/wiki) page.

### <font color="Brown">PYTHON VERSION and LIBRARIES USED:</font>

1.   Python 2.7.3, 32 bit version (required)
2.   PyWin32, build 214 (or earlier), 32 bit version (required) [See bullet 2 of the "CURRENT ISSUES" section for issues] 
3.   Matplotlib (optional, used in some of the example programs)


### <font color="Brown">CURENT ISSUES:</font>

1.   All functions (around 130) for accessing the data items that ZEMAX has made available have been implemented.  Currently, a distribution version is not available as the tool box is being updated regularly. Please download the code to a local directory in your computer and add that directory to python search path in order to use it. For more detailed instructions on using PyZDDE, please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki)

2.   Due to a known bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. This issue has been discussed and documented [here] (http://sourceforge.net/mailarchive/message.php?msg_id=28828321)


### <font color="Brown">FILES AND DIRECTORIES:</font>

*  PyZDDE                     : Top-level folder for the toolbox 
*  PyZDDE/LICENSE      : License file ([MIT License] (http://opensource.org/licenses/MIT))
*  PyZDDE/pyzdde.py   : Main file containing the PyZDDE code. The user must import this module to use the toolbox as "import pyzdde"
*  PyZDDE/zemaxoperands.py     : File containing the ZEMAX operands class and associated functions
*  PyZDDE/zemaxbuttons.py       : File containing the ZEMAX buttons class and associated functions.
*  PyZDDE/Test             : Folder containing the unit-test code for PyZDDE 
*  PyZDDE/Examples     : Folder containing example codes that uses PyZDDE
*  PyZDDE/ZMXFILES    : Folder containing a few Zemax Lens files for testing the functionality of the toolbox.


