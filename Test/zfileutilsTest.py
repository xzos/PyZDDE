#-----------------------------------------------------------------------------------------
# Name:        zfileutilsTest.py
# Purpose:     To test specific functions in the zfileutils modules. These functions are
#              meant to test specific scenarios and are generally independent of the
#              unittest functions.
#
# Licence:     MIT License
#-----------------------------------------------------------------------------------------
'''test functions in the module zfileutils
'''
from __future__ import print_function
import os as os
#from struct import unpack, pack 
import pyzdde.zfileutils as zfu
#import ctypes as _ctypes

testdir = os.path.dirname(os.path.realpath(__file__))

#%% Helper functions

def get_full_path(filename):
    """returns the full path name of the file.

    Parameters
    ----------
    filename : string
        name (with extension) of the file
    
    Assumption
    ----------    
    the files are in the 'testdata' directory under the same directory as this file
    """
    return os.path.join(testdir, 'testdata', filename)

def compare_files_nbytes(filename1, filename2, formatChar):
    """compare byte chunks in files ``filename1`` and ``filename2`` from beginning
    to end. The chunks size is determined by ``formatChar`` 
    """
    #FIXME!!! function doesn't work for formatChar = 'h'
    format_size = {'c':1, 'h':2, 'i':4, 'I':4, 'd':8, 'f':4}
    retVal = False
    i = 0    
    with open(filename1, 'rb') as f1, open(filename2, 'rb') as f2:
        while True:
            nbytes_f1 = zfu.read_n_bytes(f1, formatChar)
            nbytes_f2 = zfu.read_n_bytes(f2, formatChar)
            if nbytes_f1 and nbytes_f2:
                assert nbytes_f1 == nbytes_f2, \
                ("{} in file1 mismatached with {} in file 2 at byte offset {}."
                .format(nbytes_f1, nbytes_f2, i*format_size[formatChar]))
                i += 1
            elif nbytes_f1 and not nbytes_f2:
                print("Incomplete comparison as file1 ended. Matched upto {} bytes"
                      .format(i*format_size[formatChar]))
                break;
            elif not nbytes_f1 and nbytes_f2:
                print("Incomplete comparison as file2 ended. Matched upto {} bytes"
                      .format(i*format_size[formatChar]))
                break;
            else:
                print("Result of file compare: Files match")
                retVal = True
                break;
    return retVal
    
    
#%% ZRD file utils test
    
def test_uncompressed_zrd_read_write():
    """test the functions for reading and writing uncompressed full Data (UFD) Zemax
    ray data (ZRD) files 
    """
    print("\nTEST FOR ZRD UNCOMPRESSED FULL DATA (UFD) FILE:")
    ufdfiles = ['ColorFringes_TenRays_SplitRays_UFD.ZRD',
                'ColorFringes_TenRays_NoSplit_UFD.ZRD',
                'Beamsplitter_OneRay_SplitRays_UFD.ZRD',
                'Beamsplitter_TenRay_SplitRays_UFD.ZRD']
    for f in ufdfiles:
        zrd0 = zfu.readZRDFile(get_full_path(f))
        #print("Number of rays in the uncompressed ZRD file = ", len(zrd0))
        f2Write = 'zrdfile_test_UFD_write.ZRD'
        zfu.writeZRDFile(zrd0, get_full_path(f2Write), 'uncompressed')
        # validate the data written to the file
        zrd1 = zfu.readZRDFile(get_full_path(f2Write))
        # Compare the written file to the read file
        #print("version:", zrd0[0].zrd_version)
        assert zrd0[0].zrd_version == zrd1[0].zrd_version
        assert zrd0[0].n_segments == zrd1[0].n_segments
        compare_files_nbytes(get_full_path(f), 
                             get_full_path(f2Write), 'i')
    print("Uncompressed ZRD read/write successful.")
    
def test_compressed_zrd_read_write(): 
    """test the functions for reading and writing compressed ZRD file formats CBD 
    (compressed basic data) and CFD (compressed full data)
    """
    print("\nTEST FOR ZRD COMPRESSED BASIC DATA (CBD)")
    # Test read
    cbdfiles = ['ColorFringes_TenRays_SplitRays_CBD.ZRD',
                'ColorFringes_TenRays_NoSplitRays_CBD.ZRD']
    for f in cbdfiles:
        # Currently, readZRDFile() cannot read compressed data
        try:
            zrd = zfu.readZRDFile(get_full_path(f))
        except Exception as e:
            print('Expected exception :', e)
    # Test write
    # Nothing to test, since (currently) the writeZRDFile() only writes
    # to CFD format for compressed file type 
    
    print("\nTEST FOR ZRD COMPRESSED FILES FULL DATA (CFD)")
    # Test read
    cfdfiles = ['ColorFringes_TenRays_SplitRays_CFD.ZRD',]
    for f in cfdfiles:
        # Currently, readZRDFile() cannot read compressed data
        try:
            zrd = zfu.readZRDFile(get_full_path(f))
        except Exception as e:
            print('Expected exception :', e)
    # Test write. Since (currently) we cannot read CFD files, we will read UFD file and
    # convert to CFD file
    #ufdfiles = ['ColorFringes_TenRays_SplitRays_UFD.ZRD',]
                #'ColorFringes_TenRays_NoSplitRays_UFD.ZRD']

    #TODO!!! write test functions after the writeZRDFile() function is complete
    #for ufdf, cfdf in zip(ufdfiles, cfdfiles):
    #    ufdzrd = zfu.readZRDFile(get_full_path(ufdf))
    #    cfdf2Write = 'zrdfile_test_CFD_write.ZRD'
    #    zfu.writeZRDFile(ufdzrd, get_full_path(cfdf2Write), 'compressed')
        
    
if __name__ == '__main__':  
    test_uncompressed_zrd_read_write()
    test_compressed_zrd_read_write()