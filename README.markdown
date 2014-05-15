## PyZDDE: Python Zemax Dynamic Data Exchange

![zenodo DOI](https://zenodo.org/badge/3811/indranilsinharoy/PyZDDE.png)

##### Current revision:

0.7.6 (Last significant update on 05/11/2014)

Brief change-log is available in the [News and Updates](https://github.com/indranilsinharoy/PyZDDE/wiki/News-and-updates) page.

##### Issues and how you can help

The  [issues page](https://github.com/indranilsinharoy/PyZDDE/issues?state=open) lists current open issues, bugs, and feature request. If you find any bug, or would like to suggest any improvement to the PyZDDE library please add them to the issues page. Your ideas, feedback and suggestions are more then welcome.

Thank you very much.

##### Contributions and credits

You are encouraged to use, provide feedbacks and contribute to the PyZDDE project. The generous people who have contributed to PyZDDE are in [Contributors](https://github.com/indranilsinharoy/PyZDDE/wiki/Contributors). Thanks a lot to all of you.

Other projects that are using PyZDDE are listed [here](https://github.com/indranilsinharoy/PyZDDE/wiki/Projects-using-PyZDDE).

#### Hello world

Here is a simple but complete "Hello world" code which gets and prints the version of Zemax.   

```python
import pyzdde.zdde as pyz
link = pyz.createLink()
print("Hello Zemax version: ", link.zGetVersion())
link.close()
```

#### Overview

PyZDDE is a Python-based standalone extension for communicating with [ZEMAX] (http://www.radiantzemax.com/) using the DDE protocol. ZEMAX (now OpticStudio) is an optical design and analysis software. PyZDDE is similar to---and very much inspired by---the Matlab-based [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) developed by Derek Griffith at CSIR. At this point in time PyZDDE is not as extensive as MZDDE. 
PyZDDE can be used with regular Python scripts as well as in an interactive environment such as an IPython shell, [QtConsole] (http://ipython.org/ipython-doc/dev/interactive/qtconsole.html) or [IPython Notebook] (http://ipython.org/ipython-doc/dev/interactive/htmlnotebook.html). The ability to interact with ZEMAX from an IPython Notebook using PyZDDE can be a useful tool for teaching, recording thought process during an optical design, design documentation and design presentation/ sharing. PyZDDE is supported on both Python 2.7 and Python 3.3/3.4. It also supports both Unicode and ascii text. 


Currently, PyZDDE is a work in progress. All the functions (125 in total) for accessing the ZEMAX "data items" for extensions have been implemented. In addition, there are more than15 helper functions. The following is a sample of the extra helper functions in PyZDDE (more will be added in future):

 `zSetTimeout`, `zExecuteZPLMacro`, `zSpiralSpot`, `zGetSeidelAberration`, `zSetFieldTuple`, `zGetFieldTuple`, `zSetWaveTuple`, `zGetWaveTuple`, `zCalculateHiatus`, `zGetPupilMagnification`, `ipzCaptureWindow`, `ipzGetFirst`, `ipzGetPupil`, `ipzGetSystemAper`, `ipzGetTextWindow`.

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
import pyzdde.zdde as pyz    # import pyzdde module
ln = pyz.createLink()        # create DDE link object
ln.zPushLens(1)              # method of type 1
ln.zExecuteZPLMacro('CEN')   # method of type 2
ln.ipzCaptureWindow2('Lay')  # method of type 3
```

Helper functions of type 4 can be accessed from the `zdde` module directly.

For example:

```python
pyz.zo.findZOperand("decenter")  # method of type 4 (can also be called as pyz.findZOperand)
pyz.numAper(0.25)                # method of type 4
```

#### Is there anything missing?
The short answer is yes! PyZDDE doesn't support array/ bulk ray tracing at this point in time. I hope in the near future this feature will be implemented. May be you can help (please look in the [issues page](https://github.com/indranilsinharoy/PyZDDE/issues/21))


#### Getting started, usage, and other documentation:
Getting started with PyZDDE is really very simple as shown in the "Hello world" program above. Please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki). It has detailed guide on how to start using PyZDDE.

#### Dependencies:

1.   Python 2.7 / Python 3.3 and above; 32/64 bit version
2.   Matplotlib (optional, used in some of the example programs)

#### License:
The code is under the [MIT License] (http://opensource.org/licenses/MIT).

