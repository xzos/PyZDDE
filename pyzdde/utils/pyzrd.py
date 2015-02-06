import ctypes 
import struct as _struct


class zemax_ray():
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
        
        # Data fields of an uncompressed rayfile
        self.compressed_zrd = [('status', ctypes.c_int),
                    ('level', ctypes.c_int),
                    ('hit_object', ctypes.c_int),
                    ('parent', ctypes.c_int),
                    ('xybin', ctypes.c_int),
                    ('lmbin', ctypes.c_int),
                    ('x', ctypes.c_float),
                    ('y', ctypes.c_float),
                    ('z', ctypes.c_float),
                    ('intensity', ctypes.c_float),
                    ]

        # Data fields of a compressed rayfile
        self.uncompressed_zrd = [('status', ctypes.c_int),
                    ('level', ctypes.c_int),
                    ('hit_object', ctypes.c_int),
                    ('hit_face', ctypes.c_int),
                    ('unused', ctypes.c_int),
                    ('in_object', ctypes.c_int),
                    ('parent', ctypes.c_int),
                    ('storage', ctypes.c_int),
                    ('xybin', ctypes.c_int),
                    ('lmbin', ctypes.c_int),
                    ('index', ctypes.c_double),
                    ('starting_phase', ctypes.c_double),
                    ('x', ctypes.c_double),
                    ('y', ctypes.c_double),
                    ('z', ctypes.c_double),
                    ('l', ctypes.c_double),
                    ('m', ctypes.c_double),
                    ('n', ctypes.c_double),
                    ('nx', ctypes.c_double),
                    ('ny', ctypes.c_double),
                    ('nz', ctypes.c_double),
                    ('path_to', ctypes.c_double),
                    ('intensity', ctypes.c_double),
                    ('phase_of', ctypes.c_double),
                    ('phase_at', ctypes.c_double),
                    ('exr', ctypes.c_double),
                    ('exi', ctypes.c_double),
                    ('eyr', ctypes.c_double),
                    ('eyi', ctypes.c_double),
                    ('ezr', ctypes.c_double),
                    ('ezi', ctypes.c_double)
                    ]

        # data fields of the Zemax text ray file
        self.nsq_source_fields =  [
                    ('x', ctypes.c_double),
                    ('y', ctypes.c_double),
                    ('z', ctypes.c_double),
                    ('l', ctypes.c_double),
                    ('m', ctypes.c_double),
                    ('n', ctypes.c_double),
                    ('intensity', ctypes.c_double),
                    ('wavelength', ctypes.c_double)
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
            self.rays.append(zemax_ray())
            for field in self.rays[-1].nsq_source_fields:
                 setattr(self.rays[-1],field[0], eval(field[0]+'['+str(ii)+']'))
                 
#    def write_source_file(self,filename):
        
                
        

def readZRD(filename, file_type):
    if file_type == 'uncompressed':
        fields = 'uncompressed_zrd'
    elif file_type == 'compressed':
        fields = 'compressed_zrd'
        
    f = open(filename, "rb")
    version = _struct.unpack('i', f.read(4))[0]
    n_segments = _struct.unpack('i', f.read(4))[0]
    zrd = []
    while f.read(1):
        f.seek(-1,1)
        zrd.append(zemax_ray())
        zrd[-1].version = version
        zrd[-1].file_type = file_type
        n_segments_follow = _struct.unpack('i', f.read(4))[0]
        for ss in range(n_segments_follow):
            for field in getattr(zrd[-1],fields):
                # set the format character depending on the data type
                if field[1]== ctypes.c_int:
                    format_char = 'i'
                elif field[1]== ctypes.c_double:
                    format_char = 'd'
                elif field[1]== ctypes.c_float:
                    format_char = 'f'
                # set the value of the respective field            
                getattr(zrd[-1], field[0]).append(_struct.unpack(format_char, f.read(ctypes.sizeof(field[1])))[0])
    f.close()
    return zrd

def writeZRD(rayArray, filename,file_type):

    if file_type == 'uncompressed':
        fields = 'uncompressed_zrd'
    elif file_type == 'compressed':
        fields = 'compressed_zrd'

    f = open(filename, "wb")
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
                if field[1]== ctypes.c_int:
                    format_char = 'i'
                elif field[1]== ctypes.c_double:
                    format_char = 'd'
                elif field[1]== ctypes.c_float:
                    format_char = 'f'

                f.write(_struct.pack(format_char, getattr(rayArray[rr], field[0])[ss]))
    f.close()
                