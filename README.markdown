## PyZDDE: Python Zemax Dynamic Data Exchange

##### Current revision:

0.7.5 (Last significant update on 05/05/2014)

Brief change log is available in the [News and Updates](https://github.com/indranilsinharoy/PyZDDE/wiki/News-and-updates) page.

##### Issues and how you can help

The  [issues page](https://github.com/indranilsinharoy/PyZDDE/issues?state=open) lists current open issues, bugs, and feature request. If you find any bug, or would like to suggest any improvement to the current PyZDDE library please add them to the issues page. Your ideas, feedback and suggestions are more then welcome.

Thank you very much.

##### Contributions and credits

The generous people who have contributed to PyZDDE are in [Contributors](https://github.com/indranilsinharoy/PyZDDE/wiki/Contributors). Thanks a lot to all of you.

Other projects that are using PyZDDE are listed [here](https://github.com/indranilsinharoy/PyZDDE/wiki/Projects-using-PyZDDE).

#### Overview

PyZDDE is a standalone ZEMAX extension written in Python for communicating with [ZEMAX] (http://www.radiantzemax.com/)  using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. ZEMAX (now OpticStudio) is a leading software tool for design and analysis of optical systems. This toolbox, which implements all of the data items listed in the ZEMAX Extensions chapter of the ZEMAX manual, provides access to ZEMAX from Python. It is similar to and very much inspired by the [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) in Matlab which was developed by Derek Griffith at CSIR. However, at this point, it is not as extensive as MZDDE. PyZDDE can be used with regular Python scripts and also in an interactive environment such as an IPython shell, [QtConsole] (http://ipython.org/ipython-doc/dev/interactive/qtconsole.html) or [IPython Notebook] (http://ipython.org/ipython-doc/dev/interactive/htmlnotebook.html). The ability to interact with ZEMAX from an IPython Notebook using PyZDDE can be a useful tool for both teaching and documentation.


Currently, PyZDDE is a work in progress. All the functions (125 in total) for accessing the ZEMAX "data items" for extensions have been implemented. In addition, there are more than15 helper functions. The following is a sample of the extra helper functions in PyZDDE (more will be added in future):

 `zSetTimeout`, `zExecuteZPLMacro`, `zSpiralSpot`, `zGetSeidelAberration`, `zSetFieldTuple`, `zGetFieldTuple`, `zSetWaveTuple`, `zGetWaveTuple`, `zCalculateHiatus`, `zGetPupilMagnification`, `ipzCaptureWindow2`, `ipzGetFirst`, `ipzGetPupil`, `ipzGetSystemAper`, `ipzGetTextWindow`.

A complete list of helper functions is available [here](https://github.com/indranilsinharoy/PyZDDE/wiki/List-of-helper-functions-in-PyZDDE).

At this point in time, a distribution version is not available as the tool box is being updated regularly.

Please download the code to a local directory in your computer and add that directory to python search path in order to use it. For detailed instructions on using PyZDDE, please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki)


There are 4 types of functions in the toolbox:

1.  Functions for accessing ZEMAX using the data items defined in the "ZEMAX EXTENSIONS" chapter of the ZEMAX manual. These functions' names start with "z" and the rest of the function names matches the data item defined by Zemax. For example `zGetSolve` for the data item GetSolve, `zSetSolve` for the data item SetSolve, etc.
2.  Helper functions to enhance the toolbox functionality beyond just the data items, such as `zLensScale`, `zCalculateHiatus`, `zSpiralSpot`. Also, there are other utilities which increase the capability of the toolbox such as `zOptimize2`, `zSetWaveTuple`, `zExecuteZPLMacro`, etc. More functions are expected to be added over time.
3.  Few functions such as `ipzCaptureWindow`, `ipzGetTextWindow` can be used to embed analysis/graphic windows and text files from Zemax into an IPython Notebook or IPython QtConsole.
4.  There are several other functions which can be used independent of a running Zemax session.. Examples include `showZOperandList`, `findZOperand`, `findZButtonCode`, etc. Also, more functions are expected to be added over time.


All the functions prefixed with "z" or "ipz"  (types 1, 2 and 3) interact with Zemax directly and hence require a Zemax session to be running simultaneously. As they are instance methods of a pyzdde channel object, a pyzdde object needs to be created.

For example:

```python
import pyzdde.zdde as pyz   # import pyzdde module
ln = pyz.PyZDDE()   # create a pyzdde object
ln.zDDEInit()   # method of type 1
ln.zPushLens(1) # method of type 1
ln.zExecuteZPLMacro('CEN')  # method of type 2
ln.ipzCaptureWindow2('Lay') # method of type 3
```

Helper functions of type 4 can be accessed the the `zdde` module directly.

For example:

```python
pyz.zo.findZOperand("decenter") # method of type 4
pyz.numAper(0.25)  # method of type 4
```


#### Getting started, usage, and other documentation:
Please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki). It has detailed guide on how to start using PyZDDE.

#### Dependencies:

1.   Python 2.7 / Python 3.3 and above; 32/64 bit version
2.   Matplotlib (optional, used in some of the example programs)

#### License:
The code is under the [MIT License] (http://opensource.org/licenses/MIT).



