#-------------------------------------------------------------------------------
# Name:        pyzddeutils.py
# Purpose:     Utility functions for PyZDDE
#
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2014
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.5
#-------------------------------------------------------------------------------
from __future__ import print_function

# Try to import IPython if it is available (for notebook helper functions)
try:
    from IPython.core.display import display
except ImportError:
    #print("Couldn't import display from IPython.core.display")
    IPLoad = False
else:
    IPLoad = True
# Even if IPython is available for import, it is not necessary that the module
# is running in an IPython environment such as QtConsole or notebook.
# Use regular "print" function if not in IPython environment.
if IPLoad:
    try:
        get_ipython()
    except Exception:
        _print_mod = print   # regular print if not running on an IPython shell
    else:
        _print_mod = display # use IPython's display to prettify print
else:
    _print_mod = print    # regular print if IPython is not available.

class _prettifyCodeDesc(object):
    """Class to enable colorized Code-Description string output in IPython notebook

    In order to colorize the Code-Description, use the following idioms:

        print_mod(prettifyCodeDesc("CODE", "Description")
        or
        print_mode(pretifyText("text0", "text1" [, "text2", "color0", "color1", "color2"])
        or
        print_mode(boldifyText("text0", "text1" [, "text2", "color0", "color1", "color2"])

    If IPython environment is available, the above functions will produce colorized
    text output, using IPython's display module and if IPython is not available,
    simple print will be used.
    """
    def __init__(self,code,desc,color0='blue',color1='magenta'):
        self.code = code
        self.desc = desc
        self.color0 = color0
        self.color1 = color1
    def _repr_html_(self):
        return ("<font color='{c0}'>[</font><b><font color='{c1}'>{code}</font></b>"
                "<font color='{c0}'>]</font> {desc}"
                .format(c0=self.color0,c1=self.color1,code=self.code,desc=self.desc))
    def __repr__(self):
        return "[%s] %s" % (self.code, self.desc)

class _prettifyText(object):
    """Class to enable colorized string output in IPython notebook
    See _prettifyCodeDec for details
    """
    def __init__(self,text0,text1,text2='',color0='red',color1='green',color2='red'):
        self.text0 = text0
        self.text1 = text1
        self.text2 = text2
        self.color0 = color0
        self.color1 = color1
        self.color2 = color2
    def _repr_html_(self):
        return ("<font color='{c0}'>{text0}</font>"
                "<font color='{c1}'>{text1}</font>"
                "<font color='{c2}'>{text2}</font>"
                .format(c0=self.color0,c1=self.color1,c2=self.color2,
                        text0=self.text0,text1=self.text1,text2=self.text2))
    def __repr__(self):
        return "%s%s%s" % (self.text0, self.text1, self.text2)

class _boldifyText(object):
    """Class to enable colorized and bold string output in IPython notebook
    See _prettifyCodeDec for details
    """
    def __init__(self,text0,text1,text2='',color0='red',color1='green',color2='red'):
        self.text0 = text0
        self.text1 = text1
        self.text2 = text2
        self.color0 = color0
        self.color1 = color1
        self.color2 = color2
    def _repr_html_(self):
        return ("<font color='{c0}'><b>{text0}</font></b>"
                "<font color='{c1}'><b>{text1}</font></b>"
                "<font color='{c2}'><b>{text2}</font></b>"
                .format(c0=self.color0,c1=self.color1,c2=self.color2,
                        text0=self.text0,text1=self.text1,text2=self.text2))
    def __repr__(self):
        return "%s%s%s" % (self.text0, self.text1, self.text2)


