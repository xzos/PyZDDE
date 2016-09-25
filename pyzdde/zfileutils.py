# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        zfileutils.py
# Purpose:     File i/o support for zemax rayfiles, and other Zemax files
#
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
'''utility functions for handling various types of files used by Zemax. 
'''
from __future__ import print_function, division
import os as _os
import sys as _sys
import collections as _co
import ctypes as _ctypes
from struct import unpack as _unpack
from struct import pack as _pack
import math as _math

import pyzdde.config as _config
_global_pyver3 = _config._global_pyver3

if _global_pyver3:
   xrange = range

try:
    import numpy as _np
except ImportError:
    _global_np = False
else:
    _global_np = True

#%% ZRD read/write utilities

class ZemaxRay():
    """ Class that allows the creation and import of Zemax ray files
    
    """
    def __init__(self, parent = None):      
        self.file_type = '';
        self.status = []
        self.level = []
        self.hit_object = []
        self.hit_face = []
        self.unused = []
        self.in_object = []
        self.parent = []
        self.storage = []
        self.xybin = []
        self.lmbin = []
        self.index = []
        self.starting_phase = []
        self.x = []
        self.y = []
        self.z = []
        self.l = []
        self.m = []
        self.n = []
        self.nx = []
        self.ny = []
        self.nz = []
        self.path_to = []
        self.intensity = []
        self.phase_of = []
        self.phase_at = []
        self.exr = []
        self.exi = []
        self.eyr = []
        self.eyi = []
        self.ezr = []
        self.ezi = []
        self.wavelength = 0;
        self.zrd_type = 0
        self.zrd_version = 0
        self.n_segments = 1  # not being used
        

        self.compressed_zrd = [
                    ('status', _ctypes.c_uint),
                    ('level', _ctypes.c_int),
                    ('hit_object', _ctypes.c_int),
                    ('hit_face', _ctypes.c_int),
                    ('unused', _ctypes.c_int),
                    ('in_object', _ctypes.c_int),
                    ('parent', _ctypes.c_int),
                    ('storage', _ctypes.c_int),
                    ('xybin', _ctypes.c_int),
                    ('lmbin', _ctypes.c_int),
                    ('index', _ctypes.c_float),
                    ('starting_phase', _ctypes.c_float),
                    ('x', _ctypes.c_float),
                    ('y', _ctypes.c_float),
                    ('z', _ctypes.c_float),
                    ('l', _ctypes.c_float),
                    ('m', _ctypes.c_float),
                    ('n', _ctypes.c_float),
                    ('nx', _ctypes.c_float),
                    ('ny', _ctypes.c_float),
                    ('nz', _ctypes.c_float),
                    ('path_to', _ctypes.c_float),
                    ('intensity', _ctypes.c_float),
                    ('phase_of', _ctypes.c_float),
                    ('phase_at', _ctypes.c_float),
                    ('exr', _ctypes.c_float),
                    ('exi', _ctypes.c_float),
                    ('eyr', _ctypes.c_float),
                    ('eyi', _ctypes.c_float),
                    ('ezr', _ctypes.c_float),
                    ('ezi', _ctypes.c_float)
                    ]

        # Data fields and number format of a compressed rayfile
        self.uncompressed_zrd = [
                    ('status', _ctypes.c_uint),
                    ('level', _ctypes.c_int),
                    ('hit_object', _ctypes.c_int),
                    ('hit_face', _ctypes.c_int),
                    ('unused', _ctypes.c_int),
                    ('in_object', _ctypes.c_int),
                    ('parent', _ctypes.c_int),
                    ('storage', _ctypes.c_int),
                    ('xybin', _ctypes.c_int),
                    ('lmbin', _ctypes.c_int),
                    ('index', _ctypes.c_double),
                    ('starting_phase', _ctypes.c_double),
                    ('x', _ctypes.c_double),
                    ('y', _ctypes.c_double),
                    ('z', _ctypes.c_double),
                    ('l', _ctypes.c_double),
                    ('m', _ctypes.c_double),
                    ('n', _ctypes.c_double),
                    ('nx', _ctypes.c_double),
                    ('ny', _ctypes.c_double),
                    ('nz', _ctypes.c_double),
                    ('path_to', _ctypes.c_double),
                    ('intensity', _ctypes.c_double),
                    ('phase_of', _ctypes.c_double),
                    ('phase_at', _ctypes.c_double),
                    ('exr', _ctypes.c_double),
                    ('exi', _ctypes.c_double),
                    ('eyr', _ctypes.c_double),
                    ('eyi', _ctypes.c_double),
                    ('ezr', _ctypes.c_double),
                    ('ezi', _ctypes.c_double)
                    ]

        # data fields of the Zemax text ray file
        self.nsq_source_fields =  [
                    ('x', _ctypes.c_double),
                    ('y', _ctypes.c_double),
                    ('z', _ctypes.c_double),
                    ('l', _ctypes.c_double),
                    ('m', _ctypes.c_double),
                    ('n', _ctypes.c_double),
                    ('intensity', _ctypes.c_double),
                    ('wavelength', _ctypes.c_double)
                    ]
                    
    def __str__(self):
        if self.file_type == 'uncompressed':
            fields = 'uncompressed_zrd'
        elif self.file_type == 'compressed':
            fields = 'compressed_zrd'
        s = '';
        for field in getattr(self,fields):
            s = s + field[0] + ' = ' + str(getattr(self, field[0]))+'\n'
        return '\n' + s + '\n'

    def __repr__(self):
        return self.__str__()

class NSQSource():
    def __init__(self, parent = None):        
        self.n_layout_rays = 0
        self.n_analysis_rays = 0
        self.power = 1
        self.wavenumber = 1
        self.color = 0
        self.randomize = 0
        self.power_lumens_in_file = 0
        self.min_wavelength = 0
        self.max_wavelength = 0
        self.rays = [];

    def set_rays(self, x, y, z, l, m, n, intensity, wavelength):
        for ii in range(0,len(x)):
            self.rays.append(ZemaxRay())
            for field in self.rays[-1].nsq_source_fields:
                 setattr(self.rays[-1],field[0], eval(field[0]+'['+str(ii)+']'))
                 

def read_n_bytes(fileHandle, formatChar):
    """read n bytes from file. The number of bytes read is specified by the
    ``formatChar``
    
    Parameters
    ----------
    fileHandle : object
        file handle object
    formatChar : string
        format character -- 'c' (1), 'h' (2), 'i' (4), 'I' (4), 'd' (8), 'f' (4)
        
    Returns
    -------
    nbytes : 
    """
    nbytes = None    
    bytes2read = {'c':1, 'h':2, 'i':4, 'I':4, 'd':8, 'f':4}    
    packedBytes = fileHandle.read(bytes2read[formatChar])
    if packedBytes:
        try:
            nbytes = _unpack(formatChar, packedBytes)[0]
        except Exception as e:
            print("Reading bytes from file failed at position {}".format(fileHandle.tell()))
            print("packedBytes = {} of len {}".format(packedBytes, len(packedBytes)))
            raise e
    return nbytes 
        

def readZRDFile(file_name, max_segs_per_ray=1000):
    """import a ZRD file into an array of `ZemaxRay`. Currently, the function can only
    read uncompressed ZRD files.
    
    Usage: ``readZRD(filename [, max_segs_per_ray])``
    
    Parameters
    ----------
    file_name : string
        name of the zrd file to be imported
    max_segs_per_ray : integer
        maximum segments per ray. If the   #TODO!!! Complete this docstring 
    
    Returns
    -------
    zrd : list
        list of ``ZemaxRay`` objects. The stored parameters are defined in 
        the class ``ZemaxRay`` and depend on the file type
    
    Examples
    --------
    >>> zrd = readZRD('rays.zrd')
    """
    zrd = []
    FILE_CURR_POS = 1
    c_int, c_uint = _ctypes.c_int, _ctypes.c_uint 
    c_double, c_float  = _ctypes.c_double, _ctypes.c_float
    format_dict = {c_int:'i', c_uint:'I', c_double:'d', c_float:'f'}
    file_handle = open(file_name, "rb")
    first_int = read_n_bytes(file_handle, formatChar='i')
    zrd_type = _math.floor(first_int/10000)
    zrd_version = _math.fmod(first_int,10000) 
    max_n_segments = read_n_bytes(file_handle, formatChar='i')
    if zrd_type == 0:
        file_type = 'uncompressed'
    elif zrd_type == 2:
        raise NotImplementedError('Function cannot read CFD ZRD file format')
    else:
        raise NotImplementedError('Function cannot read CBD ZRD file format')
        
    comp_type = 'uncompressed_zrd' if (file_type == 'uncompressed') else 'compressed_zrd'
    while file_handle.read(1):
        file_handle.seek(-1, FILE_CURR_POS)
        ray = ZemaxRay()
        ray.zrd_version = zrd_version
        ray.zrd_type = zrd_type
        ray.n_segments = max_n_segments
        ray.file_type = file_type
        n_segments = read_n_bytes(file_handle, formatChar='i')
        fields = getattr(ray, comp_type)
        if n_segments > max_segs_per_ray:  
            print("n_segments ({}) > {} at byte-offset position {}. Closing file and"
                  " exiting".format(n_segments, max_segs_per_ray, file_handle.tell() - 4))
            file_handle.close()
            break;
        for _ in range(n_segments):
            for field, field_format in fields:
                format_char = format_dict[field_format]                
                # set the value of the respective field  
                getattr(ray, field).append(read_n_bytes(file_handle, format_char))
        zrd.append(ray)
    file_handle.close()
    return zrd


def writeZRDFile(rayArray, file_name, file_type):
    """write an array of `ZemaxRay` objects to a zrd file. 
    
    The uncompressed mode can only be used if all required data is available. Therefore 
    this function can be used to convert an uncompressed zrd file to a compressed file 
    but not vice versa. 
    
    Usage: ``writeZRD(rayArray, file_name, file_type)``   
    
    Parameters
    ----------
    rayArray : list 
        list of `ZemaxRay` objects to be saved to zrd file
    file_name: string
        name of the zrd file (provide full path)
    file_type: string
        type of the zrd file ('uncompressed' or 'compressed'). For compressed type,
        currently the function only writes as compressed full data (CFD) format
    
    Returns
    -------
    n/a
    
    Examples
    --------
    >>> writeZRD(rayArray, 'rays.zrd','uncompressed')
    """
    if file_type is not 'uncompressed':  # Temporary .... to remove after complete implementation
        raise NotImplementedError('Function cannot write to compressed file format')
    comp_type = 'uncompressed_zrd' if (file_type == 'uncompressed') else 'compressed_zrd'
    zrd_type = 0 if (file_type == 'uncompressed') else 20000
    c_int, c_uint = _ctypes.c_int, _ctypes.c_uint 
    c_double, c_float  = _ctypes.c_double, _ctypes.c_float
    format_dict = {c_int:'i', c_uint:'I', c_double:'d', c_float:'f'}
    file_handle = open(file_name, "wb")
    file_handle.write(_pack('i', rayArray[0].zrd_version+zrd_type))    
    file_handle.write(_pack('i', rayArray[0].n_segments)) # number of rays
    for ray in rayArray:
        file_handle.write(_pack('i', len(ray.status)))    # number of segments in the ray
        fields = getattr(ray, comp_type)
        for ss in range(len(ray.status)):
            for field, field_format in fields:
                format_char = format_dict[field_format] 
                file_handle.write(_pack(format_char, getattr(ray, field)[ss]))
    file_handle.close()

#%% Beam file read/write utilities

def readBeamFile(beamfilename):
    """Read in a Zemax Beam file

    Parameters
    ----------
    beamfilename : string
        the filename of the beam file to read

    Returns
    -------
    version : integer
        the file format version number
    n : 2-tuple, (nx, ny)
        the number of samples in the x and y directions
    ispol : boolean
        is the beam polarized?
    units : integer (0 or 1 or 2 or 3)
        the units of the beam, 0 = mm, 1 = cm, 2 = in, 3 for m
    d : 2-tuple, (dx, dy)
        the x and y grid spacing
    zposition : 2-tuple, (zpositionx, zpositiony)
        the x and y z position of the beam
    rayleigh : 2-tuple, (rayleighx, rayleighy)
        the x and y rayleigh ranges of the beam
    waist : 2-tuple, (waistx, waisty)
        the x and y waists of the beam
    lamda : double
        the wavelength of the beam
    index : double
        the index of refraction in the current medium
    receiver_eff : double
        the receiver efficiency. Zero if fiber coupling is not computed
    system_eff : double
        the system efficiency. Zero if fiber coupling is not computed.
    grid_pos : 2-tuple of lists, (x_matrix, y_matrix)
        lists of x and y positions of the grid defining the beam
    efield : 4-tuple of 2D lists, (Ex_real, Ex_imag, Ey_real, Ey_imag)
        a tuple containing two dimensional lists with the real and
        imaginary parts of the x and y polarizations of the beam

    """
    f = open(beamfilename, "rb")
    # zemax version number
    version = _unpack('i', f.read(4))[0]
    print("version: "+str(version))
    nx = _unpack('i', f.read(4))[0]
    ny = _unpack('i', f.read(4))[0]
    print("nx, ny:"+str(nx)+" "+str(ny))
    ispol = _unpack('i', f.read(4))[0]
    print("ispol: "+str(ispol))
    units = _unpack('i', f.read(4))[0]
    print("units: "+str(units))
    f.read(16)
    dx = _unpack('d', f.read(8))[0]
    dy = _unpack('d', f.read(8))[0]
    print("dx, dy: "+str(dx)+" "+str(dy))

    if version==0:
        zposition_x = _unpack('d', f.read(8))[0]
        print("zposition x: "+str(zposition_x))
        rayleigh_x = _unpack('d', f.read(8))[0]
        print("rayleigh x: "+str(rayleigh_x))
        lamda = _unpack('d', f.read(8))[0]
        print("lambda: "+str(lamda))
        #f.read(16)
        zposition_y = _unpack('d', f.read(8))[0]
        print("zposition_y: "+str(zposition_y))
        rayleigh_y = _unpack('d', f.read(8))[0]
        print("rayleigh_y: "+str(rayleigh_y))
        waist_y = _unpack('d', f.read(8))[0]
        print("waist_y: "+str(waist_y))
        waist_x=_unpack('d', f.read(8))[0]
        print("waist_x: "+str(waist_x))
        index=_unpack('d', f.read(8))[0]
        print("index: "+str(index))#f.read(64);
        receiver_eff = 0
        system_eff = 0
    if version==1:
        zposition_x = _unpack('d', f.read(8))[0]
        print("zposition x: "+str(zposition_x))
        rayleigh_x = _unpack('d', f.read(8))[0]
        print("rayleigh x: "+str(rayleigh_x))
        waist_x=_unpack('d', f.read(8))[0]
        print("waist_x: "+str(waist_x))
        #f.read(16)
        zposition_y = _unpack('d', f.read(8))[0]
        print("zposition_y: "+str(zposition_y))
        rayleigh_y = _unpack('d', f.read(8))[0]
        print("rayleigh_y: "+str(rayleigh_y))
        waist_y = _unpack('d', f.read(8))[0]
        print("waist_y: "+str(waist_y))
        lamda = _unpack('d', f.read(8))[0]
        print("lambda: "+str(lamda))
        index=_unpack('d', f.read(8))[0]
        print("index: "+str(index))
        receiver_eff=_unpack('d', f.read(8))[0]
        print("receiver efficiency: "+str(index))
        system_eff=_unpack('d', f.read(8))[0]
        print("system efficiency: "+str(index))
        f.read(64)  # 8 empty doubles

    rawx = [0 for x in range(2*nx*ny) ]
    for i in range(2*nx*ny):
        rawx[i] = _unpack('d', f.read(8))[0]
        #print(str(i)+" "+str(2*nx*ny)+" "+str(rawx[i]))
    rawy = [0 for x in range(2*nx*ny) ]
    if ispol:
        for i in range(2*nx*ny):
            rawy[i] = _unpack('d',f.read(8))[0]

    f.close()
    xc = 1+nx/2
    yc = 1+ny/2

    x_matrix = [[0 for x in xrange(nx)] for x in xrange(ny)]
    y_matrix = [[0 for x in xrange(nx)] for x in xrange(ny)]

    Ex_real = [[0 for x in xrange(nx)] for x in xrange(ny)]
    Ex_imag = [[0 for x in xrange(nx)] for x in xrange(ny)]

    Ey_real = [[0 for x in xrange(nx)] for x in xrange(ny)]
    Ey_imag = [[0 for x in xrange(nx)] for x in xrange(ny)]

    k = 0
    for j in range(ny):
        for i in range(nx):
            x_matrix[i][j] = (i-xc)*dx
            y_matrix[i][j] = (j-yc)*dy
            Ex_real[i][j] = rawx[k]
            Ex_imag[i][j] = rawx[k+1]
            if ispol:
                Ey_real[i][j] = rawy[k]
                Ey_imag[i][j] = rawy[k+1]
            k = k+2
    return (version, (nx, ny), ispol, units, (dx, dy), (zposition_x, zposition_y),
        (rayleigh_x, rayleigh_y), (waist_x, waist_y), lamda, index, receiver_eff, system_eff,
        (x_matrix, y_matrix), (Ex_real, Ex_imag, Ey_real, Ey_imag))

def writeBeamFile(beamfilename, version, n, ispol, units, d, zposition, rayleigh,
                 waist, lamda, index, receiver_eff, system_eff, efield):
    """Write a Zemax Beam file

    Parameters
    ----------
    beamfilename : string
        the filename of the beam file to read
    version : integer
        the file format version number
    n : 2-tuple, (nx, ny)
        the number of samples in the x and y directions
    ispol : boolean
        is the beam polarized?
    units : integer
        the units of the beam, 0 = mm, 1 = cm, 2 = in, 3  = m
    d : 2-tuple, (dx, dy)
        the x and y grid spacing
    zposition : 2-tuple, (zpositionx, zpositiony)
        the x and y z position of the beam
    rayleigh : 2-tuple, (rayleighx, rayleighy)
        the x and y rayleigh ranges of the beam
    waist : 2-tuple, (waistx, waisty)
        the x and y waists of the beam
    lamda : double
        the wavelength of the beam
    index : double
        the index of refraction in the current medium
    receiver_eff : double
        the receiver efficiency. Zero if fiber coupling is not computed
    system_eff : double
        the system efficiency. Zero if fiber coupling is not computed.
    efield : 4-tuple of 2D lists, (Ex_real, Ex_imag, Ey_real, Ey_imag)
        a tuple containing two dimensional lists with the real and
        imaginary parts of the x and y polarizations of the beam

    Returns
    -------
    status : integer
        0 = success; -997 = file write failure; -996 = couldn't convert
        data to integer, -995 = unexpected error.
    """
    try:
        f = open(beamfilename, "wb")
        # zemax version number
        f.write(_pack('i',version))
        f.write(_pack('i',n[0]))
        f.write(_pack('i',n[1]))
        f.write(_pack('i',ispol))
        f.write(_pack('i',units))
        # write 16 zeroes to pad out file
        f.write(_pack('4i',4,5,6,7))
        f.write(_pack('d',d[0]))
        f.write(_pack('d',d[1]))
        if version==0:
            f.write(_pack('d',zposition[0]))
            f.write(_pack('d',rayleigh[0]))
            f.write(_pack('d',lamda))
            f.write(_pack('d',zposition[1]))
            f.write(_pack('d',rayleigh[1]))
            f.write(_pack('d',waist[0]))
            f.write(_pack('d',waist[1]))
            f.write(_pack('d',index))
        if version==1:
            f.write(_pack('d',zposition[0]))
            f.write(_pack('d',rayleigh[0]))
            f.write(_pack('d',waist[0]))
            f.write(_pack('d',zposition[1]))
            f.write(_pack('d',rayleigh[1]))
            f.write(_pack('d',waist[1]))
            f.write(_pack('d',lamda))
            f.write(_pack('d',index))
            f.write(_pack('d',receiver_eff))
            f.write(_pack('d',system_eff))
            f.write(_pack('8d',1,2,3,4,5,6,7,8))

        (Ex_real, Ex_imag, Ey_real, Ey_imag) = efield

        for i in range(n[0]):
            for j in range(n[1]):
                f.write(_pack('d',Ex_real[i][j]))
                f.write(_pack('d',Ex_imag[i][j]))

        if ispol:
            for i in range(n[0]):
                for j in range(n[1]):
                    f.write(_pack('d',Ey_real[i][j]))
                    f.write(_pack('d',Ey_imag[i][j]))
        f.close()
        return 0
    except IOError as e:
        print("I/O error({0}): {1}".format(e.errno, e.strerror))
        return -997
    except ValueError:
        print("Could not convert data to an integer.")
        return -996
    except:
        print("Unexpected error:", _sys.exc_info()[0])
        return -995          

#%% Reading text files outputted by Zemax

# passing pyz object to readDetectorViewerTextFile() is hackish; however
# it is probably the best option right now. It will probably take some effort
# to move the relevant functions from the zdde module to here. 
def readDetectorViewerTextFile(pyz, textFileName, displayData=False):
    """read text file outputted from NSC detector viewer window

    Parameters
    ----------
    pyz : module object 
        the `pyzdde.zdde` module object. This is bit of a hack. See Examples  
    textFileName : string 
        full filename of the text file 
    displayData : bool 
        whether to return display data. if `False` (default) then only 
        the meta-data associated with the detector viewer window data 
        is returned 

    Return 
    ------
    dvwData : tuple 
        dvwData is a 1-tuple containing just ``dvwInfo`` (see below)
        if ``displayData`` is ``False`` (default).
        If ``displayData`` is ``True``, ``dvwData`` is a 2-tuple
        containing ``dvwInfo`` (a named tuple) and ``data``. ``data``
        is either a 2-tuple containing ``coordinate`` and ``values``
        as list elements if "Show as" is row/column cross-section, or 
        a 2D list of values otherwise.

        dvwInfo : named tuple
            surfNum : integer
                NSCG surface number
            detNum : integer 
                detector number 
            width, height : float
                width and height of the detector 
            xPix, yPix : integer
                number of pixels in x and y direction
            totHits : integer 
                total ray hits 
            peakIrr : float or None 
                peak irradiance (only available for Irradiance type of data) 
            totPow : float or None 
                total power (only available for Irradiance type of data)
            smooth : integer 
                the integer smoothing value  
            dType : string
                the "Show Data" type 
            x, y, z, tiltX, tiltY, tiltZ : float 
                the x, y, z positions and tilt values 
            posUnits : string 
                position units
            units : string 
                units  
            rowOrCol : string or None
                indicate whether the cross-section data is a row or column 
                cross-section 
            rowColNum : float or None
                the row or column number for cross-section data 
            rowColVal : float or None 
                the row or column value for cross-section data

        data : 2-tuple or 2-D list 
            if cross-section data then `data = (coordinates, values)` where, 
            `coordinates` and `values` are 1-D lists otherwise, `data` is a 
            2-D list of grid data. Note that the coherent phase data is 
            in degrees.

    Examples
    -------- 
    >>> import pyzdde.zdde as pyz 
    >>> import pyzdde.zfileutils as zfu 
    >>> info = zfu.readDetectorViewerTextFile(pyz, textFileName)
    >>> # following line assumes row/column cross-section data 
    >>> info, coordinates, values = zfu.readDetectorViewerTextFile(pyz, textFileName, True)
    >>> # following line assumes 2d data 
    >>> info, gridData = zfu.readDetectorViewerTextFile(pyz, textFileName, True)         
    """
    line_list = pyz._readLinesFromFile(pyz._openFile(textFileName))
    # Meta data
    detNumSurfNumPat = r'Detector\s*\d{1,4}\s*,\s*NSCG\sSurface\s*\d{1,4}'
    detNumSurfNum = line_list[pyz._getFirstLineOfInterest(line_list, detNumSurfNumPat)]
    detNum = int(pyz._re.search(r'\d{1,4}', detNumSurfNum.split(',')[0]).group())
    nscgSurfNum = int(pyz._re.search(r'\d{1,4}', detNumSurfNum.split(',')[1]).group())
    sizePixelsHitsPat = r'Size[0-9a-zA-Z\s\,\.]*Pixels[0-9a-zA-Z\s\,\.]*Total\sHits'
    sizePixelsHits = line_list[pyz._getFirstLineOfInterest(line_list, sizePixelsHitsPat)]
    sizeinfo, pixelsinfo , hitsinfo =  sizePixelsHits.split(',')
    #note: width<-->rows<-->xPix;; height<-->cols<-->yPix 
    width, height = [float(each) for each in pyz._re.findall(r'\d{1,4}\.\d{1,8}', sizeinfo)]
    xPix, yPix =  [int(each) for each in pyz._re.findall(r'\d{1,6}', pixelsinfo)]
    totHits = int(pyz._re.search(r'\d{1,10}', hitsinfo).group())

    #peak irradiance and total power. only present for irradiance types
    peakIrr, totPow = None, None 
    peakIrrLineNum = pyz._getFirstLineOfInterest(line_list, 'Peak\sIrradiance')
    if peakIrrLineNum:
        peakIrr = float(pyz._re.search(r'\d{1,4}\.\d{3,8}[Ee][-\+]\d{3}', 
                                       line_list[peakIrrLineNum]).group())
        totPow = float(pyz._re.search(r'\d{1,4}\.\d{3,8}[Ee][-\+]\d{3}', 
                                      line_list[peakIrrLineNum + 1]).group())

    # section of text starting with 'Smoothing' (common to all)
    smoothLineNum = pyz._getFirstLineOfInterest(line_list, 'Smoothing')
    smooth = line_list[smoothLineNum].split(':')[1].strip()
    smooth = 0 if smooth == 'None' else int(smooth)
    dType = line_list[smoothLineNum + 1].split(':')[1].strip()
    posX = float(line_list[smoothLineNum + 2].split(':')[1].strip()) # 'Detector X'
    posY = float(line_list[smoothLineNum + 3].split(':')[1].strip()) # 'Detector Y'
    posZ = float(line_list[smoothLineNum + 4].split(':')[1].strip()) # 'Detector Z'
    tiltX = float(line_list[smoothLineNum + 5].split(':')[1].strip())
    tiltY = float(line_list[smoothLineNum + 6].split(':')[1].strip())
    tiltZ = float(line_list[smoothLineNum + 7].split(':')[1].strip())
    posUnits = line_list[smoothLineNum + 8].split(':')[1].strip()
    units =  line_list[smoothLineNum + 9].split(':')[1].strip()

    # determine "showAs" type 
    rowPat = r'Row\s[0-9A-Za-z]*,\s*Y'
    colPat = r'Column\s[0-9A-Za-z]*,\s*X'
    rowColPat = '|'.join([rowPat, colPat])
    showAsRowCol = pyz._getFirstLineOfInterest(line_list, rowColPat)

    if showAsRowCol:
        # exatract specific meta data
        rowOrColPosLine = line_list[showAsRowCol]
        rowOrColIndicator = rowOrColPosLine.split(' ', 2)[0]
        assert rowOrColIndicator in ('Row', 'Column'), 'Error: Unable to determine '
        '"Row" or "Column" type for the cross section plot'
        rowOrCol = 'row' if rowOrColIndicator == 'Row' else 'col' 
        #rowColNum = line_list[showAsRowCol].split(' ', 2)[1]
        rowColNum = pyz._re.search(r'\d{1,4}|Center', rowOrColPosLine).group()
        rowColNum = 0 if rowColNum=='Center' else int(rowColNum)
        rowColVal = pyz._re.search(r'-?\d{1,4}\.\d{3,8}[Ee][-\+]\d{3}', 
                                   rowOrColPosLine).group()
        rowColVal = float(rowColVal)

        if displayData:
            dataPat = (r'\s*(-?\d{1,4}\.\d{3,8}[Ee][-\+]\d{3}\s*)' + r'{{{num}}}'
                       .format(num=2)) # coordinate, value
            start_line = pyz._getFirstLineOfInterest(line_list, dataPat)
            data_mat = pyz._get2DList(line_list, start_line, yPix)
            data_matT = pyz._transpose2Dlist(data_mat)
            coordinate = data_matT[0]
            value = data_matT[1]
    else:
        # note: it is still possible to 1d data here if `showAsRowCol` was corrupted
        rowOrCol, rowColNum, rowColVal = None, None, None # meta-data not available for 2D data
        if displayData:
            dataPat = (r'\s*\d{1,4}\s*(-?\d{1,4}\.\d{3,8}([Ee][-\+]\d{3,8})*\s*)' 
                       + r'{{{num}}}'.format(num=xPix))
            start_line = pyz._getFirstLineOfInterest(line_list, dataPat)
            gridData = pyz._get2DList(line_list, start_line, yPix, startCol=1)

    deti = _co.namedtuple('dvwInfo', ['surfNum', 'detNum', 'width', 'height',
                                      'xPix', 'yPix', 'totHits', 'peakIrr',
                                      'totPow', 'smooth', 'dType', 'x', 'y', 
                                      'z', 'tiltX', 'tiltY', 'tiltZ', 'posUnits', 
                                      'units', 'rowOrCol', 'rowColNum', 'rowColVal'])
    detInfo = deti(nscgSurfNum, detNum, width, height, xPix, yPix, totHits,
                   peakIrr, totPow, smooth, dType, posX, posY, posZ, tiltX, 
                   tiltY, tiltZ, posUnits, units, rowOrCol, rowColNum, rowColVal)

    if displayData:
        if showAsRowCol:
            return (detInfo, coordinate, value)
        else:
            return (detInfo, gridData)
    else:
        return detInfo

        
#%% Zemax surface modifier utilities

def gridSagFile(z, dzBydx, dzBydy, d2zBydxdy, nx, ny, delx, dely, unitflag=0, 
                        xdec=0, ydec=0, fname='gridsag', comment=None, fext='.DAT'):
    """generates Grid Sag ASCII file for specifying the additional sag terms of the 
    grid sag surface
    
    Parameters
    ----------
    z : ndarray
        1-dim ndarray of grid sag values
    dzBydx : ndarray
        1-dim ndarray of dz/dx values of length equal to len(z)
    dzBydy : ndarray
        1-dim ndarray of dz/dy values of length equal to len(z)
    d2zBydxdy : ndarray
        1-dim ndarray of d^2(z)/dx.dy values of length equal to len(z)
    nx : integer
        number of samples along x
    ny : integer
        number of samples along y
    unitflag : integer, optional
        0 for mm, 1 for cm, 2 for in, and 3 for meters (default=0)
    xdec : integer, optional
        decenter along x (default=0)
    ydec : integer, optional
        decenter along y (default=0)
    fname : string, optional
        filename, without extension and absolute path, of the sag file to. 
        If `None`, the name `gridsag` is used. The file is saved at:
        "C:\\Users\\%userprofile%\\Documents\\Zemax\\Objects\\Grid Files\\<fname>.xxx``
        where the exact extension is determined by ``fext``
    comment : string, optional
        top comment 
    fext : string, optional
        specifies the file extension. Use ``.DAT`` (default) for sequential
        objects, and ``.GRD`` for NSC object

    Returns
    -------
    sagfilename : string
        filename of the grid sag file along with absolute path
        
    Notes
    -----
    1. The Numpy module is required.
    2. It is assumed that the data is generated (instead of being measured), 
       therefore, the ``nodata`` field is not set for any of the data points 
       in the file, i.e. all data is valid.
    3. The file is read in Zemax using the Extra Data Editor "Import" feature.
    
    See Also
    --------
    randomGridSagFile()
    """
    fname = 'gridsag' if not fname else fname
    filename = _os.path.join(_os.path.expandvars("%userprofile%"), 'Documents',
                             'Zemax\\Objects\\Grid Files', fname + fext)
    arr = _np.hstack((z, dzBydx, dzBydy, d2zBydxdy))
    with open(filename, 'w') as f:
        f.write('! {}\n'.format(comment))
        f.write('! {} {} {} {} {} {} {}  <-- First data line\n'
                .format('nx', 'ny', 'delx', 'dely', 'unitflag', 'xdec', 'ydec'))
        f.write('! {} {} {} {} <-- Other data lines\n'
                .format('z', 'dz/dx', 'dz/dy', 'd2z/dxdy'))
        f.write('{:d} {:d} {: 12.9E} {: 12.9E} {:d} {: 12.9E} {: 12.9E}\n'
                .format(nx, ny, delx, dely, unitflag, xdec, ydec))
        for row in arr:
            f.write('{: 12.9E} {: 12.9E} {: 12.9E} {: 12.9E}\n'.format(*row))
    return filename
    
def randomGridSagFile(mu=0, sigma=1, semidia=1, nx=201, ny=201, unitflag=0, 
                         xdec=0, ydec=0, fname='gridsag_randn', comment=None, fext='.DAT'):
    """generates grid sag ASCII file with Gaussian distributed sag profile
    
    Parameters
    ----------
    mu : float, optional
        mean of the normal distribution
    sigma : float, optional
        standard deviation of the normal distribution. If `sigma` is `np.inf`
        all height is set to zero (0)
    nx : integer, optional
        number of samples along x
    ny : integer, optional
        number of samples along y
    unitflag : integer, optional
        0 for mm, 1 for cm, 2 for in, and 3 for meters
    xdec : integer, optional
        decenter along x
    ydec : integer, optional
        decenter along y
    semidia : float, optional
        semi-diameter of the grid sag surface (note that grid is rectangular) 
    fname : string, optional
        filename, without extension and absolute path, of the sag file to. 
        If `None`, the name `gridsag_randn` is used. The file is saved at:
        "C:\\Users\\%userprofile%\\Documents\\Zemax\\Objects\\Grid Files\\<fname>.xxx"
    fext : string, optional
        specifies the file extension. Use ``.DAT`` (default) for sequential
        objects, and ``.GRD`` for NSC object
    comment : string, optional
        top comment 
        
    Returns
    -------
    sag : ndarray
        sag profile
    fname : string
        full file name of the ASCII file
    
    Notes
    -----
    1. The Numpy module is required.
    2. The function doesn't compute any of the derivatives. Therefore Zemax 
       computes them automatically.
    3. The file is read in Zemax using the Extra Data Editor "Import" feature.
    4. The interpolation method for the grid sag surface must be linear 
       (Use value of 1 for Parameter 0 of Grid sag surface).
    5. The wave propagation method must use the angular spectrum method
       (option available under the POP tab in the LDE by right clicking on 
       the surface).
       
    See Also
    --------
    gridSagFile(),
    """
    delx = 2.0*semidia/(nx-1)
    dely = 2.0*semidia/(ny-1)
    
    if sigma is not _np.inf:
        sag = _np.random.normal(loc=mu, scale=sigma, size=(nx, ny))
    else:
        sag = _np.zeros((nx, ny))
    
    dzBydx =  _np.zeros((nx*ny, 1))
    dzBydy = dzBydx.copy()
    d2zBydxdy = dzBydx.copy()
    z=sag.reshape(nx*ny, 1)
    
    gridsagfile = gridSagFile(z, dzBydx, dzBydy, d2zBydxdy, nx, ny, 
                              delx, dely, unitflag, xdec, ydec, 
                              fname, comment, fext)
    return z, gridsagfile
