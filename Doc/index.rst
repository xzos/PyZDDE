PyZDDE: Python Zemax DDE Extension
==================================

Installing
----------

To install **PyZDDE** use a command prompt and ``cd`` to the extracted folder. Then execute 

.. code:: python
  
     python setup.py install


Things to do after installation
-------------------------------


Initial setup
~~~~~~~~~~~~~

1. **PUSH LENS PERMISSION:** All operations through the DDE affect the lens in the DDE server (except for array ray tracing). In order to copy the lens from the DDE server to the Zemax application /LDE, you need to "push" the lens from the server to the LDE. To do so, please enable the option "Allow Extensions to Push Lenses", under File->Preferences->Editors tab.
2. **ANSI/UNICODE TEXT ENCODING:** PyZDDE supports both ANSI and UNICODE text from Zemax. Please set the appropriate text encoding in PyZDDE by calling module function `pyz.setTextEncoding(text_encoding)` (assuming that PyZDDE was imported as `import pyzdde.zdde as pyz`). By default, UNICODE text encoding is set in PyZDDE. You can check the current text encoding by calling `pyz.getTextEncoding()` function. Please note that you need to do this only when you change the text setting in Zemax and not for every session.
3. **PURE NSC MODE:** (This is more of a note) If want to work on an optical design in pure NSC mode, please start ZEMAX in pure NSC mode before initiating the communication with PyZDDE. There is no way to switch the ZEMAX mode using external interfaces.

ZPL macros files supplied with PyZDDE
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PyZDDE comes with few ZPL macro files that are present in the directory "ZPLMacros". They are occasionally used by PyZDDE (for example in the function ``ipzCaptureWindowLQ()``). Please copy/move the files from the folder "ZPLMacros" to the folder where Zemax/ Optic studio expects to find ZPL macros (By default, this folder is ``C:\<username>\Documents\ZEMAX\Macros``). A copy of the "ZPLMacros" folder is always available in (installed with) the PyZDDE package.


Examples
~~~~~~~~

Examples shipped with PyZDDE are in the folder "Examples". Please move the examples to your desired location after extracting the PyZDDE package. 


Usage
-----

Start Zemax/ OpticStudio, import PyZDDE in a script/ IPython notebook, create a DDE communication channel, control/ communicate/ ray-trace, etc in Zemax, and close the link. For example, the following is a simple ``Hello world`` program.

.. code:: python

    import pyzdde.zdde as pyz
    ln = pyz.createLink()
    print("Hello Zemax version: ", ln.zGetVersion())
    ln.close()


Modules in PyZDDE
-----------------

-  **zdde** (``import pyzdde.zdde as pyz``): The main module in PyZDDE that provides all dataitems related functions for interacting with Zemax/OpticStudio using the DDE interface.
-  **arraytrace** (``import pyzdde.arraytrace as at``): provides functions for tracing large number of rays
-  **zfileutils** (``import pyzdde.zfileutils as zfu``): provides helper functions for various Zemax file handling operations such as reading and writing beam files, .ZRD files, creating .DAT and .GRD files for grid phase /grid sag surfaces, etc.
-  **systems** (``import pyzdde.systems as osys``): provides helper functions for quickly creating basic optical systems.


Getting started, usage, and other documentation
-----------------------------------------------

Getting started with PyZDDE is really very simple as shown in the "Hello world" program above. Please refer to the [GitHub wiki page](https://github.com/indranilsinharoy/PyZDDE/wiki). It has detailed guide on how to start using PyZDDE. The wiki also has various other documentation on several topics. Please refer to the GitHub wiki documentation. Also, examples included with PyZDDE are viewable [here](http://nbviewer.ipython.org/github/indranilsinharoy/PyZDDE/tree/master/Examples/).