#-------------------------------------------------------------------------------
# Name:        pyzrd.py
# Purpose:     File i/o support for zemax rayfiles
#
# Copyright:   (c) Florian Hudelist
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.1
#-------------------------------------------------------------------------------


import ctypes as _ctypes
import struct as _struct


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
        self.n_segments = 1
        

        # Data fields and number format of an compressed (full) rayfile
        self.compressed_full_zrd = [('status', _ctypes.c_int),
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
        self.uncompressed_zrd = [('status', _ctypes.c_int),
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
    zrd : An array of zemax rays. The stored parameters are defined in the ZemaxRay class and depend on the file type
    
    Examples
    --------
    zrd = readZRD('rays.zrd','uncompressed')
    
    """
        
    f = open(file_name, "rb")
    version = _struct.unpack('i', f.read(4))[0]
    
    if version < 10000:
        file_type = 'uncompressed'
        fields = 'uncompressed_zrd'
    elif version > 20000:
        print('Ray file data format is Compressed Full Data (CFD) which cannot be imported')
        return -1
    else:
        print('Ray file data format is Compressed Basic Data (CBD) which cannot be imported')
        return -1
           

    n_segments = _struct.unpack('i', f.read(4))[0]
    zrd = []
    while f.read(1):
        f.seek(-1,1)
        zrd.append(ZemaxRay())
        zrd[-1].version = version        
        zrd[-1].file_type = file_type
        n_segments_follow = _struct.unpack('i', f.read(4))[0]
        for ss in range(n_segments_follow):
            for field in getattr(zrd[-1],fields):
                # set the format character depending on the data type
                if field[1]== _ctypes.c_int:
                    format_char = 'i'
                elif field[1]== _ctypes.c_double:
                    format_char = 'd'
                elif field[1]== _ctypes.c_float:
                    format_char = 'f'
                # set the value of the respective field            
                getattr(zrd[-1], field[0]).append(_struct.unpack(format_char, f.read(_ctypes.sizeof(field[1])))[0])
    f.close()
    return zrd

def writeZRDFile(rayArray, file_name,file_type):
    """ 
    writeZRD(rayArray, file_name, file_type)
    
    Write an array of ZemaxRay() to a zrd file. The uncompressed mode can only be used if all required data is available. Therefore this function can be used to convert an uncompressed zrd file to a compressed file but not vice versa.
    
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
    f.write(_struct.pack('i',rayArray[0].version))
    # number of rays in the ray array
    f.write(_struct.pack('i',len(rayArray)))
    for rr in range(0,len(rayArray)):
        # number of surfaces in the ray
        f.write(_struct.pack('i',len(rayArray[rr].status)))
        for ss in range(0,len(rayArray[rr].status)):
            for field in getattr(rayArray[rr],fields):
                # set the format character depending on the data type
                if field[1]== _ctypes.c_int:
                    format_char = 'i'
                elif field[1]== _ctypes.c_double:
                    format_char = 'd'
                elif field[1]== _ctypes.c_float:
                    format_char = 'f'

                f.write(_struct.pack(format_char, getattr(rayArray[rr], field[0])[ss]))
    f.close()
                