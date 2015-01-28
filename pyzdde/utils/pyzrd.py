import ctypes 
import struct as _struct
import numpy


class zemax_ray():
    def __init__(self, parent = None):        
        self.status = numpy.asarray([])
        self.level = numpy.asarray([])
        self.hit_object = numpy.asarray([])
        self.hit_face = numpy.asarray([])
        self.unused = numpy.asarray([])
        self.in_object = numpy.asarray([])
        self.parent = numpy.asarray([])
        self.storage = numpy.asarray([])
        self.xybin = numpy.asarray([])
        self.lmbin = numpy.asarray([])
        self.index = numpy.asarray([])
        self.starting_phase = numpy.asarray([])
        self.x = numpy.asarray([])
        self.y = numpy.asarray([])
        self.z = numpy.asarray([])
        self.l = numpy.asarray([])
        self.m = numpy.asarray([])
        self.n = numpy.asarray([])
        self.nx = numpy.asarray([])
        self.ny = numpy.asarray([])
        self.nz = numpy.asarray([])
        self.path_to = numpy.asarray([])
        self.intensity = numpy.asarray([])
        self.phase_of = numpy.asarray([])
        self.phase_at = numpy.asarray([])
        self.exr = numpy.asarray([])
        self.exi = numpy.asarray([])
        self.eyr = numpy.asarray([])
        self.eyi = numpy.asarray([])
        self.ezr = numpy.asarray([])
        self.ezi = numpy.asarray([])
        self.version = 0
        self.n_segments = 1

        self.data_fields = [('status', ctypes.c_int),
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
                    
    def __str__(self):
        s = '';
        for field in self.data_fields:
            s = s + field[0] + ' = ' + str(getattr(self, field[0]))+'\n'
        return s

    def __repr__(self):
        return self.__str__


def readZRD(filename):
    f = open(filename, "rb")
    version = _struct.unpack('i', f.read(4))[0]
    n_segments = _struct.unpack('i', f.read(4))[0]
    zrd = []
    while f.read(1):
        f.seek(-1,1)
        zrd.append(zemax_ray())
        zrd[-1].version = version
        n_segments_follow = _struct.unpack('i', f.read(4))[0]
        for ss in range(n_segments_follow):
            for field in zrd[-1].data_fields:
                setattr(zrd[-1],field[0], numpy.append(getattr(zrd[-1], field[0]), _struct.unpack(('d' if field[1]==ctypes.c_double else 'i'), f.read(ctypes.sizeof(field[1])))[0]))
    f.close()
    return zrd

def writeZRD(rayArray, filename):
    f = open(filename, "wb")
    # zemax version number
    f.write(_struct.pack('i',rayArray[0].version))
    # number of rays in the ray array
    f.write(_struct.pack('i',len(rayArray)))
    for rr in numpy.arange(len(rayArray)):
        # number of surfaces in the ray
        f.write(_struct.pack('i',len(rayArray[rr].status)))
        for ss in numpy.arange(len(rayArray[rr].status)):
            for field in rayArray[rr].data_fields:
                f.write(_struct.pack(('d' if field[1]==ctypes.c_double else 'i'),getattr(rayArray[rr], field[0])[ss]))

                