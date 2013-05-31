## PyZDDE: Python Zemax Dynamic Data Exchange

#### Overview

PyZDDE is a toolbox for communicating with [ZEMAX] (http://www.radiantzemax.com/)  using the Microsoft's Dynamic Data Exchange (DDE) messaging protocol. ZEMAX is a leading software tool for design and analysis of optical systems. This toolbox, which implements all of the data items listed in the ZEMAX Extensions chapter of the ZEMAX manual, provides access and control to ZEMAX from Python. It is similar to and very much inspired by the [MZDDE toolbox] (http://kb-en.radiantzemax.com/KnowledgebaseArticle50204.aspx) in Matlab which is developed by Derek Griffith at CSIR. However, at this point, it is not as extensive as MZDDE.

#### Usage and getting started:
Please refer to the [Wiki] (https://github.com/indranilsinharoy/PyZDDE/wiki) page.

#### Dependencies:

1.   Python 2.7 and above, 32 bit version (required)
2.   PyWin32, build 214 (or earlier), 32 bit version (required) [See bullet 2 of the "CURRENT ISSUES" section for issues] 
3.   Matplotlib (optional, used in some of the example programs)


#### Current issues:

1.   All functions (around 130) for accessing the data items that ZEMAX have been implemented. Currently, a distribution version is not available as the tool box is being updated regularly. Please download the code to a local directory in your computer and add that directory to python search path in order to use it. For more detailed instructions on using PyZDDE, please refer to the [Wiki page] (https://github.com/indranilsinharoy/PyZDDE/wiki)

2.   Due to a known bug in the PyWin32 library, please use a 32-bit version, built 214 (or earlier) of PyWin32. This issue has been discussed and documented [here] (http://sourceforge.net/mailarchive/message.php?msg_id=28828321)


#### License:
[MIT License] (http://opensource.org/licenses/MIT)



