#-------------------------------------------------------------------------------
# Name:        setup.py
# Purpose:     Standard module installation script
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.1
#-------------------------------------------------------------------------------
"""
This script will install the PyZDDE library into your Lib/site-packages directory
as a standard Python module. It can then be imported like any other module package.

To use:
- Open the command prompt (cmd.exe)
- CD to the directory where this file is located
- issue the command
   python setup.py install
"""

from setuptools import setup, find_packages

setup(
    name="PyZDDE",
    version="1.1.00",
    description='Zemax / OpticStudio standalone extension using Python',
    author='Indranil Sinharoy',
    author_email='indranil_leo@yahoo.com',
    license='MIT',
    keywords='zemax opticstudio extensions dde',
    url='https://github.com/indranilsinharoy/PyZDDE',
    packages=find_packages()
)
