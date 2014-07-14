#-------------------------------------------------------------------------------
# Name:        config.py
# Purpose:     Configuration module
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2014
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.8.0
#-------------------------------------------------------------------------------
from os import path as _path
import sys as _sys

try: # Python 3.x
   from configparser import SafeConfigParser
except ImportError: # Python 2.x
    from ConfigParser import SafeConfigParser

# Helper functions for chaning settings.ini file
def setTextEncoding(txt_encoding=0):
    """sets the text encoding to match the TXT encoding in Zemax

    Parameters
    ----------
    txt_encoding (integer) : 0 = ASCII
                             1 = UNICODE

    Returns
    -------
    status : True if success; False if fail
    """
    global _global_use_unicode_text
    status = False
    if txt_encoding: # unicode
        success = changeEncodingConfiguration(encoding_type=0, encoding=1)
        if success:
            new_encoding = getEncodingConfiguration(encoding_type=0)
            if new_encoding == 'unicode':
                _global_use_unicode_text = True
                status = True
    else: # ascii
        success = changeEncodingConfiguration(encoding_type=0, encoding=0)
        if success:
            new_encoding = getEncodingConfiguration(encoding_type=0)
            if new_encoding == 'ascii':
                _global_use_unicode_text = False
                status = True
    return status

def getTextEncoding():
    """returns the current text encoding set in PyZDDE """
    return getEncodingConfiguration(encoding_type=0)

def getSettingsFileFullName():
    """returns the full path of the settings.ini file"""
    settings_file_name = 'settings.ini'
    dirpath = _path.dirname(_path.realpath(__file__))
    filepath = _path.join(dirpath, settings_file_name)
    return filepath

def getEncodingConfiguration(encoding_type=0):
    """get the encoding in the configuration file

    Parameters
    ----------
    encoding_type (integer) : 0 = Text encoding; 1 = File encoding

    Return
    ------
    encoding (string) : 'ascii' or 'unicode'
    """
    parser = SafeConfigParser()
    parser.read(getSettingsFileFullName())
    if encoding_type: # file encoding
        pass
    else: # text encoding
        encoding = parser.get('zemax_text_encoding', 'encoding')
    return encoding

def changeEncodingConfiguration(encoding_type=0, encoding=0):
    """change the encoding in the configuration file (settings.ini)

    Parameters
    ----------
    encoding_type (integer) : 0 = Text encoding; 1 = File encoding
    encoding (integer) : 0 = ascii; 1 = unicode

    Return
    ------
    success_status (bool) : True = success; False = failed
    """
    parser = SafeConfigParser()
    parser.read(getSettingsFileFullName())
    status = True
    if encoding_type:  # file encoding
        pass
    else:  # text encoding
        if encoding:  # unicode
            parser.set(section='zemax_text_encoding',
                       option='encoding',
                       value='unicode')
        else:  # ascii
            parser.set(section='zemax_text_encoding',
                       option='encoding',
                       value='ascii')
    try:
        cfgfile = open(getSettingsFileFullName(), 'w')
        parser.write(cfgfile)
        cfgfile.close()
    except:
        status = False
    return status

# Get Zemax text encoding variable from settings .ini file
text_encoding = getEncodingConfiguration(encoding_type=0)
if text_encoding == 'unicode':
    _global_use_unicode_text = True
    #print('config.py:: text setting is UNICODE')
else:
    _global_use_unicode_text = False
    #print('config.py:: text setting is ASCII')


# Python version variable
_global_pyver3 = None
# Check Python version and set the global version variable
if _sys.version_info[0] < 3:
    _global_pyver3 = False
else:
    _global_pyver3 = True


