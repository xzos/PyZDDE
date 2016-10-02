..  image:: https://raw.githubusercontent.com/indranilsinharoy/PyZDDE/master/Doc/Images/logo_text_small.png


Python Zemax Dynamic Data Exchange
-----------------------------------

|DOI|

Current revision
'''''''''''''''''

2.0.3 (Last significant update on 10/02/2016)



Change log
~~~~~~~~~~
Brief change-log is available in the `News and
Updates <https://github.com/indranilsinharoy/PyZDDE/wiki/08.-News-and-updates>`__
page.


Examples
''''''''

Examples included with PyZDDE are in the folder "Examples". Please move the examples to your desired location after extracting the PyZDDE package. 


Hello world
~~~~~~~~~~~

Here is a simple but complete "Hello world" code which prints the version of Zemax. (If you are using Python 2.x, don't forget to add
``from __future__ import print_function`` before these lines.)

.. code:: python

    import pyzdde.zdde as pyz
    ln = pyz.createLink() # DDE link object
    print("Hello Zemax version: ", ln.zGetVersion())
    ln.close()

More examples (view online)
^^^^^^^^^^^^^^^^^^^^^^^^^^^

A gallery of notebooks demonstrating the use of PyZDDE within Jupyter (previously IPython) notebooks 
are `here <https://github.com/indranilsinharoy/PyZDDE/wiki/03.-Using-PyZDDE-in-Jupyter:-A-Gallery-of-notebooks>`__.

Examples of using Zemax interactively from a Python shell is `here <https://github.com/indranilsinharoy/PyZDDE/wiki/02.-Using-PyZDDE-interactively-in-a-Python-shell>`_.

Example Python scripts are
`here <https://github.com/indranilsinharoy/PyZDDE/tree/master/Examples/Scripts/>`__.

Examples specific to array ray tracing are catalogued
`here <https://github.com/indranilsinharoy/PyZDDE/wiki/05.-Examples-of-array-ray-tracing>`__.

In addition, the repository
`Intro2LensDesignByGeary <https://github.com/indranilsinharoy/Intro2LensDesignByGeary>`__
contains notes from few chapters of the book "Introduction to Lens
Design," by Joseph M. Geary, in the form of IPython notebooks.


Install PyZDDE from PyPI
''''''''''''''''''''''''

You can either use `pip <https://pip.pypa.io/en/stable/>`__ or the ``setup.py`` script 
from the extracted folder.

Use the following command from the command line to install PyZDDE from PyPI:

.. code:: python

  pip install pyzdde


Note 1. By default only the latest stable version is installed by default, using
the above command. To install pre-release versions add the 
`--pre <https://pip.pypa.io/en/latest/reference/pip_install.html#pre-release-versions>`__
flag:

.. code:: python
  
  pip install pyzdde --pre  


Note 2. When you install PyZDDE using pip (the above method), the "Examples" folder 
will not be downloaded. However, pip is the most convenient way to install Python packages.
Also, ensure that you have the `pip` package in your environment. 

If you would like to download the "Examples" folder too, please download and extract PyZDDE 
package from the `Python Package Index <https://pypi.python.org/pypi/PyZDDE>`__. Then,  
``cd`` into the extracted folder where the ``setup.py`` script is visible and execute 
the following in a command prompt:

.. code:: python

  python setup.py install

If you would like to see what files were added and where you may use:

.. code:: python

  python setup.py install --record files.txt

A list of all files that were added and their locations will be available in the 
file "files.txt" in the same directory.  

Note 3. To uninstall pyzdde using pip use

.. code:: python

  pip uninstall pyzdde


Get the latest code
'''''''''''''''''''

To get the latest PyZDDE code please download / fork / clone from 
`GitHub repository <https://github.com/indranilsinharoy/PyZDDE>`__.


Documentation
'''''''''''''

The PyZDDE documentation is currently hosted in the `GitHub Wiki <https://github.com/indranilsinharoy/PyZDDE/wiki>`__.


Initial setup
'''''''''''''

1. **PUSH LENS PERMISSION:** All operations through the DDE affect the lens in the DDE server (except for array ray tracing). In order to copy the lens from the DDE server to the Zemax application /LDE, you need to "push" the lens from the server to the LDE. To do so, please enable the option "Allow Extensions to Push Lenses", under File->Preferences->Editors tab.
2. **ANSI/UNICODE TEXT ENCODING:** PyZDDE supports both ANSI and UNICODE text from Zemax. Please set the appropriate text encoding in PyZDDE by calling module function `pyz.setTextEncoding(text_encoding)` (assuming that PyZDDE was imported as `import pyzdde.zdde as pyz`). By default, UNICODE text encoding is set in PyZDDE. You can check the current text encoding by calling `pyz.getTextEncoding()` function. Please note that you need to do this only when you change the text setting in Zemax and not for every session.
3. **PURE NSC MODE:** (This is more of a note) If want to work on an optical design in pure NSC mode, please start ZEMAX in pure NSC mode before initiating the communication with PyZDDE. There is no way to switch the ZEMAX mode using external interfaces.


**ZPL macros files supplied with PyZDDE**

PyZDDE comes with few ZPL macro files that are present in the directory "ZPLMacros". They are occasionally used by PyZDDE (for example in the function ``ipzCaptureWindowLQ()``). Please copy/move the files from the folder "ZPLMacros" to the folder where Zemax/ Optic studio expects to find ZPL macros (By default, this folder is ``C:\<username>\Documents\ZEMAX\Macros``). A copy of the "ZPLMacros" folder is always available in (installed with) the PyZDDE package.


Modules in PyZDDE
'''''''''''''''''

-  **zdde** (``import pyzdde.zdde as pyz``): The main module in PyZDDE that provides all dataitems related functions for interacting with Zemax/OpticStudio using the DDE interface.
-  **arraytrace** (``import pyzdde.arraytrace as at``): provides functions for tracing large number of rays
-  **zfileutils** (``import pyzdde.zfileutils as zfu``): provides helper functions for various Zemax file handling operations such as reading and writing beam files, .ZRD files, creating .DAT and .GRD files for grid phase /grid sag surfaces, etc.
-  **systems** (``import pyzdde.systems as osys``): provides helper functions for quickly creating basic optical systems.
-  **misc** (``import pyzdde.misc as mys``): contains miscellaneous collection of utility functions that may be used with PyZDDE.

Features
~~~~~~~~

-  Functions for using all "data items" defined in Zemax manual
-  Supports both Python 2.7 and Python 3.3/3.4
-  Supports both Unicode and extended ascii text
-  Over 80 additional functions for more efficient use (more will be added in future). Examples include ``zSetTimeout()``,
   ``zExecuteZPLMacro()``, ``zGetSeidelAberration()``, ``zSetFieldTuple()``,
   ``zGetFieldTuple()``, ``zSetWaveTuple()``, ``zGetWaveTuple()``, ``zCalculateHiatus()``, ``zGetPupilMagnification()``, ``zGetPOP()``,
   ``zSetPOPSettings()``, ``zModifyPOPSettings()``, ``zGetPSF()``, ``zGetPSFCrossSec()``, ``zGetMTF()``, ``zGetImageSimulation()``.
   A list of the additional functions are available `here <https://github.com/indranilsinharoy/PyZDDE/wiki/07.-List-of-helper-functions-in-PyZDDE>`__.
-  Special functions for better interactive use with IPython notebooks.
   Examples include ``ipzCaptureWindow()``, ``ipzGetFirst()``, ``ipzGetPupil()``, ``ipzGetSystemAper()``, ``ipzGetTextWindow()``
-  Quick generation of few simple optical systems (see ``pyzdde.systems`` module)
-  Array ray tracing using a separate and standalone module ``arraytrace`` along with helper functions for performing array ray tracing.

Overview
~~~~~~~~

PyZDDE is a Python-based extension for communicating with `ZEMAX/OpticStudio <http://www.zemax.com/>`__ using the DDE
protocol. It is similar to---and very much inspired by---the Matlab-based `MZDDE toolbox <http://www.zemax.com/support/resource-center/knowledgebase/how-to-talk-to-zemax-from-matlab>`__ developed by Derek Griffith at CSIR.

PyZDDE can be used with regular Python scripts as well as in an interactive environment such as an IPython shell, 
`QtConsole <http://ipython.readthedocs.org/en/stable/interactive/qtconsole.html>`__ or `IPython Notebook <http://ipython.org/ipython-doc/dev/interactive/htmlnotebook.html>`__.

There are 4 types of functions, and a separate module for array ray tracing in the toolbox:

1. Functions for accessing ZEMAX using the data items defined in the "ZEMAX EXTENSIONS" chapter of the ZEMAX manual. These functions'
   names start with "z" and the rest of the function names matches the data item defined by Zemax. For example ``zGetSolve()`` for the data
   item "GetSolve", ``zSetSolve()`` for the data item "SetSolve", etc.
2. Helper functions to enhance the toolbox functionality beyond just the data items, such as ``zCalculateHiatus``, ``zSpiralSpot``. Also,
   there are other utilities which increase the capability of the toolbox such as ``zOptimize2()``, ``zSetWaveTuple()``,
   ``zExecuteZPLMacro()``, etc.
3. Few functions such as ``ipzCaptureWindow()``, ``ipzGetTextWindow()`` can be used to embed analysis/graphic windows and text files from
   Zemax into an IPython Notebook or IPython QtConsole.
4. There are several other functions which can be used independent of a running Zemax session. Examples include ``showZOperandList()``,
   ``findZOperand()``, ``findZButtonCode()``, etc.
5. A separate and standalone module ``arraytrace`` for performing array ray tracing.

All the functions prefixed with "z" or "ipz" (types 1, 2 and 3) interact with Zemax directly and hence require a Zemax session to be running
simultaneously. As they are instance methods of a pyzdde channel object, a pyzdde object needs to be created.

For example:

.. code:: python

    import pyzdde.zdde as pyz    # import pyzdde module
    ln = pyz.createLink()        # create DDE link object
    ln.zPushLens(1)              # method of type 1
    ln.zExecuteZPLMacro('CEN')   # method of type 2
    ln.ipzCaptureWindow2('Lay')  # method of type 3

Helper functions of type 4 can be accessed from the ``zdde`` module directly.

For example

.. code:: python

    pyz.zo.findZOperand("decenter")  # method of type 4 (same as pyz.findZOperand)
    pyz.numAper(0.25)                # method of type 4

A complete list of helper functions is available
`here <https://github.com/indranilsinharoy/PyZDDE/wiki/List-of-helper-functions-in-PyZDDE>`__.
(Please be mindful that the currently this page is not updated at the same rate at which functions are getting added)


Getting started, usage, and other documentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Getting started with PyZDDE is really very simple as shown in the "Hello world" program above. Please refer to the `Wiki page <https://github.com/indranilsinharoy/PyZDDE/wiki>`__. It has detailed guide on how to start using PyZDDE.

Dependencies
''''''''''''

The core PyZDDE library only depends on the standard Python Library. 

1. Python 2.7 / Python 3.3 and above; 32/64 bit version
2. Matplotlib (optional, used in some of the example programs)

License
'''''''

The code is under the `MIT License <http://opensource.org/licenses/MIT>`__.


Contributions and credits
'''''''''''''''''''''''''

You are encouraged to use, provide feedbacks and contribute to the PyZDDE project. The generous people who have contributed to PyZDDE are
in `Contributors <https://github.com/indranilsinharoy/PyZDDE/wiki/09.-Contributors>`__. Thanks a lot to all of you.

Other projects that are using PyZDDE are listed `here <https://github.com/indranilsinharoy/PyZDDE/wiki/10.-Projects-using-PyZDDE>`__.


Citing
''''''

If you use PyZDDE for research work, please consider citing it. Various
citation styles for PyZDDE are available from
`zenodo <https://zenodo.org/record/15763?ln=en>`__.

Chat room
''''''''''

|Gitter chat|

.. |DOI| image:: https://zenodo.org/badge/doi/10.5281/zenodo.44295.svg
   :target: http://dx.doi.org/10.5281/zenodo.44295
.. |Gitter chat| image:: https://badges.gitter.im/indranilsinharoy/PyZDDE.png
   :target: https://gitter.im/indranilsinharoy/PyZDDE
