#-------------------------------------------------------------------------------
# Name:        zfileutils.py
# Purpose:     File i/o support for zemax rayfiles
#
# Copyright:   (c) Florian Hudelist
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.1
#-------------------------------------------------------------------------------
from __future__ import print_function
import ctypes as _ctypes
from struct import unpack as _unpack
from struct import pack as _pack


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
        self.version = 0
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
                 
#    def write_source_file(self,filename):
        
                
        

def readZRDFile(file_name):
    """ 
    readZRD(filename, file_type)
    
    Import a zrd file to an array of ZemaxRay()
    
    Parameters
    ----------
    file_name: [string] name of the zrd file to be imported
    file_type: [string] type of the zrd file ('uncompressed' of 'compressed')
    
    Returns
    -------
    zrd : An array of zemax rays. The stored parameters are defined in the ZemaxRay 
        class and depend on the file type
    
    Examples
    --------
    zrd = readZRD('rays.zrd','uncompressed')
    
    """
        
    f = open(file_name, "rb")
    version = _unpack('i', f.read(4))[0]
    n_segments = _unpack('i', f.read(4))[0]
    if version < 10000:
        file_type = 'uncompressed'
        fields = 'uncompressed_zrd'
    elif version > 20000:
        print('Ray file data format is Compressed Full Data (CFD) which cannot be imported')
        return -1
    else:
        print('Ray file data format is Compressed Basic Data (CBD) which cannot be imported')
        return -1
           

    zrd = []
    while f.read(1):
        f.seek(-1,1)
        zrd.append(ZemaxRay())
        zrd[-1].version = version
        zrd[-1].n_segments = n_segments
        zrd[-1].file_type = file_type
        n_segments_follow = _unpack('i', f.read(4))[0]
        for ss in range(n_segments_follow):
            for field in getattr(zrd[-1],fields):
                # set the format character depending on the data type
                if field[1]== _ctypes.c_int:
                    format_char = 'i'
                if field[1]== _ctypes.c_uint:
                    format_char = 'I'
                elif field[1]== _ctypes.c_double:
                    format_char = 'd'
                elif field[1]== _ctypes.c_float:
                    format_char = 'f'
                # set the value of the respective field            
                getattr(zrd[-1], field[0]).append(_unpack(format_char, f.read(_ctypes.sizeof(field[1])))[0])
    f.close()
    return zrd

def read_n_bytes(fileHandle, formatChar):
    """read n bytes from file. The number of bytes read is specified by the
    ``formatChar``
    
    fileHandle : file handle
    formatChar : 'c' (1), 'h' (2), 'i' (4), 'I' (4), 'd' (8), 'f' (4)
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

def readZRDFile(file_name, file_type, max_segs_per_ray=1000):
    """ 
    readZRD(filename, file_type)
    
    Import a zrd file to an array of ZemaxRay()
    
    Parameters
    ----------
    file_name : string
        name of the zrd file to be imported
    file_type : string
        type of the zrd file ('uncompressed' of 'compressed')
    max_segs_per_ray : integer
        maximum segments per ray. If the 
    
    Returns
    -------
    zrd : list
        list of ``ZemaxRay`` objects. The stored parameters are defined in 
        the class ``ZemaxRay`` and depend on the file type
    
    Examples
    --------
    >>> zrd = readZRD('rays.zrd','uncompressed')
    
    """
    comp_type = 'uncompressed_zrd' if (file_type == 'uncompressed') else 'compressed_zrd'
    zrd = []
    FILE_CURR_POS = 1
    c_int, c_uint = _ctypes.c_int, _ctypes.c_uint 
    c_double, c_float  = _ctypes.c_double, _ctypes.c_float
    format_dict = {c_int:'i', c_uint:'I', c_double:'d', c_float:'f'}
    file_handle = open(file_name, "rb")
    version = read_n_bytes(file_handle, formatChar='i')
    max_n_segments = read_n_bytes(file_handle, formatChar='i')
    while file_handle.read(1):
        file_handle.seek(-1, FILE_CURR_POS)
        ray = ZemaxRay()
        ray.version = version
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

def _writeZRDFile(rayArray, file_name,file_type):
    """ 
    writeZRD(rayArray, file_name, file_type)
    
    Write an array of ZemaxRay() to a zrd file. The uncompressed mode can only be used 
    if all required data is available. Therefore this function can be used to convert an 
    uncompressed zrd file to a compressed file but not vice versa.
    
    Parameters
    ----------
    rayArray : [ZemaxRay()] list of ZemaxRay elements to be saved
    file_name: [string] name of the zrd file to be imported
    file_type: [string] type of the zrd file ('uncompressed' of 'compressed')
    
    Returns
    -------
    n/a
    
    Examples
    --------
    writeZRD(rayArray, 'rays.zrd','uncompressed')
    
    """

    if file_type == 'uncompressed':
        fields = 'uncompressed_zrd'
    elif file_type == 'compressed':
        fields = 'compressed_zrd'

    f = open(file_name, "wb")
    # zemax version number
    f.write(_pack('i',rayArray[0].version))
    # number of rays in the ray array
    f.write(_pack('i',len(rayArray)))
    for rr in range(0,len(rayArray)):
        # number of surfaces in the ray
        f.write(_pack('i',len(rayArray[rr].status)))
        for ss in range(0,len(rayArray[rr].status)):
            for field in getattr(rayArray[rr],fields):
                # set the format character depending on the data type
                if field[1]== _ctypes.c_int:
                    format_char = 'i'
                elif field[1] == _ctypes.c_uint:
                    format_char = 'I'
                elif field[1]== _ctypes.c_double:
                    format_char = 'd'
                elif field[1]== _ctypes.c_float:
                    format_char = 'f'

                f.write(_pack(format_char, getattr(rayArray[rr], field[0])[ss]))
    f.close()

def writeZRDFile(rayArray, file_name, file_type):
    """ 
    writeZRD(rayArray, file_name, file_type)
    
    Write an array of ZemaxRay() to a zrd file. The uncompressed mode can only be used if 
    all required data is available. Therefore this function can be used to convert an 
    uncompressed zrd file to a compressed file but not vice versa.
    
    Parameters
    ----------
    rayArray : [ZemaxRay()] list of ZemaxRay elements to be saved
    file_name: [string] name of the zrd file to be imported
    file_type: [string] type of the zrd file ('uncompressed' of 'compressed')
    
    Returns
    -------
    n/a
    
    Examples
    --------
    writeZRD(rayArray, 'rays.zrd','uncompressed')
    
    """
    comp_type = 'uncompressed_zrd' if (file_type == 'uncompressed') else 'compressed_zrd'
    c_int, c_uint = _ctypes.c_int, _ctypes.c_uint 
    c_double, c_float  = _ctypes.c_double, _ctypes.c_float
    format_dict = {c_int:'i', c_uint:'I', c_double:'d', c_float:'f'}
    file_handle = open(file_name, "wb")
    file_handle.write(_pack('i', rayArray[0].version))    
    file_handle.write(_pack('i', rayArray[0].n_segments)) # number of rays
    for ray in rayArray:
        file_handle.write(_pack('i', len(ray.status)))    # number of segments in the ray
        fields = getattr(ray, comp_type)
        for ss in range(len(ray.status)):
            for field, field_format in fields:
                format_char = format_dict[field_format] 
                file_handle.write(_pack(format_char, getattr(ray, field)[ss]))
    file_handle.close()
                