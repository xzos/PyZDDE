## PyZDDE: Python Zemax Dynamic Data Exchange

#### Overview

PyZDDE is a toolbox, which is written in Python, is used for communicating with [ZEMAX] (http://www.radiantzemax.com/)  using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. ZEMAX is a leading software tool for design and analysis of optical systems. This toolbox, which implements all of the data items listed in the ZEMAX Extensions chapter of the ZEMAX manual, provides access to ZEMAX from Python. It is similar to and very much inspired by the [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) in Matlab which was developed by Derek Griffith at CSIR. However, at this point, it is not as extensive as MZDDE. PyZDDE can be used with regular Python scripts and also in an interactive environment such as an IPython shell, [QtConsole] (http://ipython.org/ipython-doc/dev/interactive/qtconsole.html) or [IPython Notebook] (http://ipython.org/ipython-doc/dev/interactive/htmlnotebook.html). The ability to interact with ZEMAX from an IPython Notebook using PyZDDE can be a useful tool for both teaching and documentation.


Currently, PyZDDE is a work in progress. All the functions (126 in total) for accessing the ZEMAX "data items" for extensions have been implemented. In addition, there are around 10 more helper functions. At this point in time, a distribution version is not available as the tool box is being updated regularly. Please download the code to a local directory in your computer and add that directory to python search path in order to use it. For more detailed instructions on using PyZDDE, please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki)


There are 4 types of functions in the toolbox:

1.  Functions for accessing ZEMAX using the data items defined in the "ZEMAX EXTENSIONS" chapter of the ZEMAX manual. These functions' names start with "z" and the rest of the function names matches the data item defined by Zemax. For example `zGetSolve` for the data item GetSolve, `zSetSolve` for the data item SetSolve, etc. 
2.  Helper functions to enhance the toolbox functionality beyond just the data items, such as `zLensScale`, `zCalculateHiatus`, `zSpiralSpot`. Also, there are other utilities which increase the capability of the toolbox such as `zOptimize2`, `zSetWaveTuple`, `zExecuteZPLMacro`, etc. More functions are expected to be added over time.
3.  Few functions such as `ipzCaptureWindow`, `ipzGetTextWindow` can be used to embed analysis/graphic windows and text files from Zemax into an IPython Notebook or IPython QtConsole. 
4.  There are several other functions that do not require to interact with Zemax directly. Examples include `showZOperandList`, `findZOperand`, `findZButtonCode`, etc. These functions may be used even without Zemax. Also, more functions are expected to be added over time.


All the functions prefixed with "z" or "ipz" interact with Zemax directly and hence require Zemax to running simultaneously.



#### Usage and getting started:
Please refer to the [Wiki] (https://github.com/indranilsinharoy/PyZDDE/wiki) page.

#### Dependencies:

1.   Python 2.7 and above, 32 bit version (required)
2.   PyWin32, build 214 (or earlier), 32 bit version (required) [See bullet 2 of the "Current issues" section for issues] 
3.   Matplotlib (optional, used in some of the example programs)


#### Current issues:

1.   Due to a known bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. This issue has been discussed and documented [here] (http://sourceforge.net/mailarchive/message.php?msg_id=28828321)
If you are using any of the Python Scientific Package distributions such as Enthought EPD, Enthough Canopy, Continumm Analytics' Anaconda, Python(x,y) or WinPython, please ensure that you are using version 214 of PyWin32. At this point in time (June 2013), EPD and Canopy comes with the 214 version of PyWin32. So, Anaconda, Python(x,y) and WinPython users will have to roll-back to version 214.


#### License:
[MIT License] (http://opensource.org/licenses/MIT)



