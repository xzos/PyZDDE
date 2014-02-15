## PyZDDE: Python Zemax Dynamic Data Exchange

#### Overview

PyZDDE, which is a toolbox written in Python, is used for communicating with [ZEMAX] (http://www.radiantzemax.com/)  using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. ZEMAX is a leading software tool for design and analysis of optical systems. This toolbox, which implements all of the data items listed in the ZEMAX Extensions chapter of the ZEMAX manual, provides access to ZEMAX from Python. It is similar to and very much inspired by the [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) in Matlab which was developed by Derek Griffith at CSIR. However, at this point, it is not as extensive as MZDDE. PyZDDE can be used with regular Python scripts and also in an interactive environment such as an IPython shell, [QtConsole] (http://ipython.org/ipython-doc/dev/interactive/qtconsole.html) or [IPython Notebook] (http://ipython.org/ipython-doc/dev/interactive/htmlnotebook.html). The ability to interact with ZEMAX from an IPython Notebook using PyZDDE can be a useful tool for both teaching and documentation.


Currently, PyZDDE is a work in progress. All the functions (126 in total) for accessing the ZEMAX "data items" for extensions have been implemented. In addition, there are around 10 more helper functions. At this point in time, a distribution version is not available as the tool box is being updated regularly. Please download the code to a local directory in your computer and add that directory to python search path in order to use it. For more detailed instructions on using PyZDDE, please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki)


There are 4 types of functions in the toolbox:

1.  Functions for accessing ZEMAX using the data items defined in the "ZEMAX EXTENSIONS" chapter of the ZEMAX manual. These functions' names start with "z" and the rest of the function names matches the data item defined by Zemax. For example `zGetSolve` for the data item GetSolve, `zSetSolve` for the data item SetSolve, etc. 
2.  Helper functions to enhance the toolbox functionality beyond just the data items, such as `zLensScale`, `zCalculateHiatus`, `zSpiralSpot`. Also, there are other utilities which increase the capability of the toolbox such as `zOptimize2`, `zSetWaveTuple`, `zExecuteZPLMacro`, etc. More functions are expected to be added over time.
3.  Few functions such as `ipzCaptureWindow`, `ipzGetTextWindow` can be used to embed analysis/graphic windows and text files from Zemax into an IPython Notebook or IPython QtConsole. 
4.  There are several other functions which can be used independent of a running Zemax session.. Examples include `showZOperandList`, `findZOperand`, `findZButtonCode`, etc. Also, more functions are expected to be added over time.


All the functions prefixed with "z" or "ipz"  (types 1, 2 and 3) interact with Zemax directly and hence require a Zemax session to be running simultaneously. As they are instance methods of a pyzdde channel object, a pyzdde object needs to be created.

For example:

```python
import pyzdde.zdde as pyz        # import pyzdde module
ln = pyz.PyZDDE()                      # create a pyzdde object
ln.zDDEInit()                                # method of type 1
ln.zPushLens(1)                          # method of type 1
ln.zExecuteZPLMacro('CEN')   # method of type 2
ln.ipzCaptureWindow2('Lay')    # method of type 3
```

Helper functions of type 4 can be accessed the the `zdde` module directly. 

For example:

```python
pyz.zo.findZOperand("decenter")   # method of type 4
pyz.numAper(U)                                # method of type 4
```


#### Usage and getting started:
Please refer to the [Wiki] (https://github.com/indranilsinharoy/PyZDDE/wiki) page. It has detailed guide on how to start using PyZDDE.

#### Dependencies:

1.   Python 2.7 and above, 32/64 bit version (haven't tested on Python 3.x)
2.   PyWin32, build 218.4 or later
3.   Matplotlib (optional, used in some of the example programs)


#### Current issues:

1.   ~~Due to a known bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. This issue has been discussed and documented [here] (http://sourceforge.net/mailarchive/message.php?msg_id=28828321)
If you are using any of the Python Scientific Package distributions such as Enthought EPD, Enthough Canopy, Continumm Analytics' Anaconda, Python(x,y) or WinPython, please ensure that you are using version 214 of PyWin32. At this point in time (June 2013), EPD and Canopy comes with the 214 version of PyWin32. So, Anaconda, Python(x,y) and WinPython users will have to roll-back to version 214.~~

UPDATE:
It seems that the DDE communication error in PyWin32 has been resolved in both 32-bit and 64-bit versions of built 218.4 (I have verified this using both [WinPython](http://winpython.sourceforge.net/) and [Anaconda](https://store.continuum.io/cshop/anaconda/) distributions). So PyZDDE can be used with both 32 and 64 bit versions of Python. If you are getting an `ImportError` during `pyzdde.zdde` import, then it is most likely due to the PyWin32. Please ensure that you are using the appropriate version of PyWin32.


#### License:
[MIT License] (http://opensource.org/licenses/MIT)



