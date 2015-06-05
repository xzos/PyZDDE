# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        config.py
# Purpose:     Configuration module
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
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
    txt_encoding : integer
        0 = ASCII; 1 = UNICODE

    Returns
    -------
    status : bool
        ``True`` if success; ``False`` if fail
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
    """returns the current text encoding set in PyZDDE
    This is an internal helper function
    """
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
    encoding_type : integer
        0 = Text encoding; 1 = File encoding

    Returns
    -------
    encoding : string
        'ascii' or 'unicode'
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
    encoding_type : integer
        0 = Text encoding; 1 = File encoding
    encoding : integer
        0 = ascii; 1 = unicode

    Returns
    -------
    status : bool
        True = success; False = failed
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

def getImageMagickSettings():
    """return the use-flag and imageMagick installation directory
    settings

    Parameters
    ----------
    None

    Returns
    -------
    use_flag : bool
        if ``True``, then PyZDDE uses the installed version of ImageMagick
        software. If ``False``, then the version of ImageMagick that
        comes with PyZDDE will be used.
    imageMagick_dir : string
        ImageMagick installation directory.
    """
    parser = SafeConfigParser()
    parser.read(getSettingsFileFullName())
    image_magick_settings = parser.items('imageMagick_config')
    use_flag_index, dir_index, value_index = 0, 1, 1
    use_flag = eval(image_magick_settings[use_flag_index][value_index])
    imageMagick_dir = image_magick_settings[dir_index][value_index]
    return use_flag, imageMagick_dir


def setImageMagickSettings(use_installed_ImageMagick,
                          imageMagick_dir=None):
    """set the use-flag and imageMagick installation directory settings

    Parameters
    ----------
    use_installed_ImageMagick : bool
        boolean flag to indicate whether to use installed version
        of ImageMagick (``True``) or not (``False``)
    imageMagick_dir : string, optional
        full path to the installation directory. For example:
        ``C:\\Program Files\\ImageMagick-6.8.9-Q8``

    Returns
    -------
    status : bool
        ``True`` = success; ``False`` = fail
    """
    status = True
    parser = SafeConfigParser()
    parser.read(getSettingsFileFullName())
    parser.set(section='imageMagick_config',
               option='useInstalled',
               value=str(use_installed_ImageMagick))
    if imageMagick_dir:
        parser.set(section='imageMagick_config',
                   option='imgMagickIinstallDir',
                   value=imageMagick_dir)
    try:
        cfgfile = open(getSettingsFileFullName(), 'w')
        parser.write(cfgfile)
        cfgfile.close()
    except:
        status = False
    return status


# -------------------------------------------------------
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