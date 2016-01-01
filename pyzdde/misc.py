# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        misc.py
# Purpose:     Miscellaneous utility functions.
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
'''miscellaneous utility functions. 
'''
from __future__ import print_function, division
import os as _os 
import zdde as _pyz
import collections as _co

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




def _draw_plane(ln, space='img', dist=0, surfName=None, semiDia=None):
    """function to draw planes at the points specified by `dist` in the 
    space specified by `space`
    
    Parameters
    ----------
    ln : pyzdde object
        active link object
    space : string (`img` or `obj`), optional
        image space or object space in which the plane is specified. 'img' for 
        image space, 'obj' for object space. This info is required because 
        Zemax returns distances that are measured w.r.t. surface 1 (@LDE) in 
        object space, and w.r.t. IMG in image space. See the Assumptions.
    dist : float, optional
        distance along the optical axis of the plane from surface 2 (@LDE) if 
        `space` is `obj` else from the IMG surface. This assumes that surface 1
        is a dummy surface
    surfName : string, optional
        name to identify the surf in the LDE, added to the comments column
    semiDia : real, optional
        semi-diameter of the surface to set 
        
    Returns
    -------
    None
    
    Assumptions (important to read)
    -------------------------------
    The function assumes (for the purpose of this study) that surface 1 @ LDE is 
    a dummy surface at certain distance preceding the first actual lens surface. 
    This enables the rays entering the lens to be visible in the Zemax layout 
    plots even if the object is at infinity. So the function inserts the planes 
    (and their associated dummy surfaces) beginning at surface 2.
    """
    numSurf = ln.zGetNumSurf()
    inSurfPos = numSurf if space=='img' else 2 # assuming that the first surface will be a dummy surface
    ln.zInsertDummySurface(surfNum=inSurfPos, thick=dist, semidia=0, comment='dummy')    
    ln.zInsertSurface(inSurfPos+1)
    ln.zSetSurfaceData(inSurfPos+1, ln.SDAT_COMMENT, surfName)
    if semiDia:
        ln.zSetSemiDiameter(surfNum=inSurfPos+1, value=semiDia)
    thickSolve, pickupSolve = 1, 5
    frmSurf, scale, offset, col = inSurfPos, -1, 0, 0
    ln.zSetSolve(inSurfPos+1, thickSolve, pickupSolve, frmSurf, scale, offset, col)
    
def gaussian_lens_formula(u=None, v=None, f=None, infinity=10e20):
    """return the third value of the Gaussian lens formula, given any two

    Parameters
    ----------
    u : float, optional
        object distance from first principal plane. 
    v : float, optional
        image distance from rear principal plane 
    f : float, optional
        focal length
    infinity : float
        numerical value to represent infinity (default=10e20)

    Returns
    -------
    glfParams : namedtuple
        named tuple containing the Gaussian Lens Formula parameters

    Notes
    ----- 
    Both object and image distances are considered positive.   

    Examples
    --------
    >>> gaussian_lens_formula(u=30, v=None, f=10)
    glfParams(u=30, v=15.0, f=10)
    >>> gaussian_lens_formula(u=30, v=15)
    glfParams(u=30, v=15, f=10.0)
    >>> gaussian_lens_formula(u=1e20, f=10)
    glfParams(u=1e+20, v=10.0, f=10)
    """
    glfParams = _co.namedtuple('glfParams', ['u', 'v', 'f'])
    def unknown_distance(knownDistance, f):
        try: 
            unknownDistance = (knownDistance * f)/(knownDistance - f)
        except ZeroDivisionError:
            unknownDistance = infinity 
        return unknownDistance

    def unknown_f(u, v):
        return (u*v)/(u+v)

    if sum(i is None for i in [u, v, f]) > 1:
        raise ValueError('At most only one parameter can be None')

    if f is None:
        if not u or not v:
            raise ValueError('f cannot be determined from input')
        else:
            f = unknown_f(u, v)
    else:
        if u is None:
            u = unknown_distance(v, f)
        else:
            v = unknown_distance(u, f)
    return glfParams(u, v, f)

def get_cardinal_points(ln):
    """Returns the distances of the cardinal points (along the optical axis).
    
    For multiple wavelengths, the distances are averaged.
    
    Parameters
    ----------
    ln : object
        PyZDDE object
    
    Returns
    -------
    fpObj : float
        distance of object side focal plane from surface # 1 in the LDE, 
        irrespective of which surface is defined as the global reference  
    fpImg : float
        distance of image side focal plane from IMG surface
    ppObj : float
        distance of the object side principal plane from surface # 1 in the 
        LDE, irrespective of which surface is defined as the global 
        reference surface 
    ppImg : float
        distance of the image side principal plane from IMG
    
    Notes
    -----
    1. The data is consistant with the cardinal data in the Prescription file
       in which, the object side data is with respect to the first surface in the LDE. 
    2. If there are more than one wavelength, then the distances are averaged.  
    """
    zmxdir = _os.path.split(ln.zGetFile())[0]
    textFileName = _os.path.join(zmxdir, "tmp.txt") 
    ln.zGetTextFile(textFileName, 'Pre', "None", 0)
    line_list = _pyz._readLinesFromFile(_pyz._openFile(textFileName))
    ppObj, ppImg, fpObj, fpImg = 0.0, 0.0, 0.0, 0.0
    count = 0
    for line_num, line in enumerate(line_list):
        # Extract the Focal plane distances
        if "Focal Planes" in line:
            fpObj += float(line.split()[3])
            fpImg += float(line.split()[4])
        # Extract the Principal plane distances.
        if "Principal Planes" in line and "Anti" not in line:
            ppObj += float(line.split()[3])
            ppImg += float(line.split()[4])
            count +=1  #Increment (wavelength) counter for averaging
    # Calculate the average (for all wavelengths) of the principal plane distances
    # This is only there for extracting a single point ... ideally the design
    # should have just one wavelength define!
    if count > 0:
        fpObj = fpObj/count
        fpImg = fpImg/count
        ppObj = ppObj/count
        ppImg = ppImg/count
    # Delete the temporary file
    _pyz._deleteFile(textFileName)
    cardinals = _co.namedtuple('cardinals', ['Fo', 'Fi', 'Ho', 'Hi'])
    return cardinals(fpObj, fpImg, ppObj, ppImg)
      
def draw_pupil_cardinal_planes(ln, firstDummySurfOff=10, cardinalSemiDia=1.2, push=True):
    """Insert paraxial pupil and cardinal planes surfaces in the LDE for rendering in
    layout plots. This is a semi-automated process; see notes.
    
    Parameters
    ----------
    ln : object
        pyzdde object
    firstDummySurfOff : float, optional 
        the thickness of the first dummy surface. This first dummy surface is 
        inserted by this function. See Notes.
    cardinalSemiDia : float, optional 
        semidiameter of the cardinal surfaces. (Default=1.2) 
    push : bool
        push lens in the DDE server to the LDE
        
    Assumptions
    -----------
    The function assumes that the lens is already focused appropriately,
    for either finite or infinite conjugate imaging. 
    
    Notes
    -----
    1. 'first dummy surface' is a dummy surface in LDE position 1 (between the 
        OBJ and the actual first lens surface) whose function is show the input 
        rays to the left of the first optical surface.
    2. The cardinal and pupil planes are drawn using standard surfaces in the LDE. 
       To ensure that the ray-tracing engine does not treat these surfaces as real 
       surfaces, we need to instruct Zemax to "ignore" rays to these surfaces. 
       Unfortunately, we cannot do it programmatically. So, after the planes have 
       been drawn, we need to manually do the following:
           1. 2D Layout settings
               a. Set number of rays to 1 or as needed
           2. For the pupil (ENPP and EXPP) and cardinal surfaces (H, H', F, F'), 
              and the dummy surfaces (except for the dummy surface named "dummy 2 
              c rays" go to "Surface Properties" >> Draw tab
               a. Select "Skip rays to this surface" 
           3. Set field points to be symmetric about the optical axis
    3. For clarity, the semi-diameters of the dummy sufaces are set to zero.
    """
    ln.zSetWave(0, 1, 1)
    ln.zSetWave(1, 0.55, 1)
    # insert dummy surface at 1 for showing the input ray
    ln.zRemoveVariables()
    # before inserting surface check to see if the object is at finite 
    # distance. If the object is at finite distance, inserting a dummy 
    # surface with finite thickness will change the image plane distance.
    # so first decrease the thickness of the object surface by the 
    # thickness of the dummy surface
    objDist = ln.zGetSurfaceData(surfNum=0, code=ln.SDAT_THICK)
    assert firstDummySurfOff < objDist, ("dummy surf. thick ({}) must be < "
                                         "than obj dist ({})!".format(firstDummySurfOff, objDist))
    if objDist < 1.0E+10:
        ln.zSetSurfaceData(surfNum=0, code=ln.SDAT_THICK, value=objDist - firstDummySurfOff)
    ln.zInsertDummySurface(surfNum=1, thick=firstDummySurfOff, semidia=0, comment='dummy 2 c rays')
    ln.zGetUpdate()
    # Draw Exit and Entrance pupil planes
    print("Textual information about the planes:\n")
    expp = ln.zGetPupil().EXPP
    print("Exit pupil distance from IMG:", expp)
    _draw_plane(ln, 'img', expp, "EXPP")
    
    enpp = ln.zGetPupil().ENPP
    print("Entrance pupil from Surf 1 @ LDE:", enpp)
    _draw_plane(ln, 'obj', enpp - firstDummySurfOff, "ENPP")

    # Get and draw the Principal planes
    fpObj, fpImg, ppObj, ppImg = get_cardinal_points(ln)

    print("Focal plane obj F from surf 1 @ LDE: ", fpObj, "\nFocal plane img F' from IMA: ", fpImg)
    _draw_plane(ln,'img', fpImg, "F'", cardinalSemiDia)
    _draw_plane(ln,'obj', fpObj - firstDummySurfOff, "F", cardinalSemiDia)
    
    print("Principal plane obj H from surf 1 @ LDE: ", ppObj, "\nPrincipal plane img H' from IMA: ", ppImg)
    _draw_plane(ln,'img', ppImg, "H'", cardinalSemiDia)
    _draw_plane(ln,'obj', ppObj - firstDummySurfOff, "H", cardinalSemiDia)

    # Check the validity of the distances
    ppObjToEnpp = ppObj - enpp
    ppImgToExpp = ppImg - expp
    focal = ln.zGetFirst().EFL
    print("Focal length: ", focal)
    print("Principal plane H to ENPP: ", ppObjToEnpp)
    print("Principal plane H' to EXPP: ", ppImgToExpp)
    v = gaussian_lens_formula(u=ppObjToEnpp, v=None, f=focal).v
    print("Principal plane H' to EXPP (abs.) "
          "calc. using lens equ.: ", abs(v))
    ppObjTofpObj = ppObj - fpObj
    ppImgTofpImg = ppImg - fpImg
    print("Principal plane H' to rear focal plane: ", ppObjTofpObj)
    print("Principal plane H to front focal plane: ", ppImgTofpImg)
    print(("""\nCheck "Skip rays to this surface" under "Draw Tab" of the """
           """surface property for the dummy and cardinal plane surfaces. """
           """See Docstring Notes for details."""))
    if push:
        ln.zPushLens(1)

