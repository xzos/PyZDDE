#-------------------------------------------------------------------------------
# Name:        systems.py
# Purpose:     Simple optical systems for quick setup with PyZDDE.
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2014
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.8.0
#-------------------------------------------------------------------------------
"""simple optical systems for quick setup with PyZDDE & Zemax. The docstring
examples assume that PyZDDe is imported as ``import pyzdde.zdde as pyz``,
a PyZDDE communication object is then created as ``ln = pyz.createLink()``
or ``ln = pyz.PyZDDE(); ln.zDDEInit()`` and ``systems`` (this module) is
imported as ``import pyzdde.systems as optsys``
"""
from __future__ import division
from __future__ import print_function

def zMakeIdealThinLens(ddeLn, fl=50, fn=5, stop_pos=0, stop_shift=0, opd_mode=1,
                       zmx_mode=0):
    """Creates an ideal thin lens of the given specification consisting of
    a STOP and a PARAXIAL surface in the zemax server.

    Parameters
    ----------
    ddeLn : object
        pyzdde object
    fl : float, optional
        focal length (measured in air of unity index) in lens units
    fn : float, optional
        f-number (image space f/#)
    stop_pos : integer (0/1), optional
        use 0 to place STOP before (to the left of) the paraxial
        surface, 1 to place STOP after (to the right)
    stop_shift : integer, optional
        axial distance between STOP and paraxial surface
    opd_mode : integer (0, 1, 2 or 3), optional
        the OPD mode, which indicates how Zemax should calculate the
        optical path difference for rays refracted by the paraxial lens.

        * ``opd_mode=0`` is fast and accurate if the aberrations are not
          severe. Zemax uses parabasal ray tracing in this mode.
        * ``opd_mode=1`` is the most accurate.
        * ``opd_mode=2`` assumes that the lens is used at infinite
          conjugates.
        * ``opd_mode=3`` is similar to ``opd_mode=0``, except that zemax
          traces paraxial rays instead of paraxial rays.

    zmx_mode : integer (0, 1, or 2), optional
        zemax mode. 0 for sequential, 1 for hybrid, 2 for mixed. Currently
        ignored.

    Returns
    -------
    None

    Examples
    --------
    >>> import pyzdde.zdde as pyz
    >>> ln = pyz.createLink()
    >>> optsys.zMakeIdealThinLens(ln)
    >>> optsys.zMakeIdealThinLens(ln, fl=100, fn=5, opd_mode=0)
    >>> optsys.zMakeIdealThinLens(ln, stop_shift=5)
    >>> optsys.zMakeIdealThinLens(ln, stop_pos=1, stop_shift=5)

    Notes
    -----
    1. For more information see "Paraxial" under "Sequential surface type
       definitions" in the Zemax manual.
    2. Use ``ln.zPushLens(1)`` update lens into the LDE
    """
    if  stop_pos < 0 or stop_pos > 1:
        raise ValueError("Expecting stop_pos to be either 0 or 1")
    epd = fl/fn
    stop_surf = 2 if stop_pos else 1
    para_surf = 1 if stop_pos else 2
    ddeLn.zNewLens()
    ddeLn.zInsertSurface(para_surf)
    ddeLn.zSetSystemAper(aType=0, stopSurf=stop_surf, apertureValue=epd)
    if stop_pos:
        ddeLn.zSetSurfaceData(para_surf, code=3, value=stop_shift)
    else:
        ddeLn.zSetSurfaceData(stop_surf, code=3, value=stop_shift)
    ddeLn.zSetSurfaceData(para_surf, code=0, value='PARAXIAL')
    ddeLn.zSetSurfaceParameter(para_surf, parameter=1, value=fl) # focallength
    ddeLn.zSetSurfaceParameter(para_surf, parameter=2, value=opd_mode)
    surf_beforeIMA, thickness_code = 2, 1
    solve_type, height, pupil_zone = 2, 0, 0 # Marginal ray height
    ddeLn.zSetSolve(surf_beforeIMA, thickness_code, solve_type, height, pupil_zone)

def zMakeIdealCollimator(ddeLn, fl=50, fn=5, ima_dist=10, opd_mode=1, zmx_mode=0):
    """Creates a collimator using an ideal thin lens of the given
    specification.

    The model consists of just 3 surfaces in the LDE -- OBJ, STOP
    (paraxial surface) and IMA plane.

    Parameters
    ----------
    ddeLn : object
        pyzdde object
    fl : float, optional
        focal length (measured in air of unity index) in lens units
    fn : float, optional
        f-number (image space f/#)
    ima_dist : integer, optional
        axial distance the paraxial surface and IMA (observation plane)
    opd_mode : integer (0, 1, 2 or 3), optional
        the OPD mode, which indicates how Zemax should calculate the
        optical path difference for rays refracted by the paraxial lens.

        * ``opd_mode=0`` is fast and accurate if the aberrations are not
          severe. Zemax uses parabasal ray tracing in this mode.
        * ``opd_mode=1`` is the most accurate.
        * ``opd_mode=2`` assumes that the lens is used at infinite
          conjugates.
        * ``opd_mode=3`` is similar to ``opd_mode=0``, except that zemax
          traces paraxial rays instead of paraxial rays.

    zmx_mode : integer (0, 1, or 2), optional
        zemax mode. 0 for sequential, 1 for hybrid, 2 for mixed. Currently
        ignored.

    Returns
    -------
    None

    Examples
    --------
    >>> optsys.zMakeIdealCollimator(ln)
    >>> optsys.zMakeIdealCollimator(ln, fl=100, fn=5, opd_mode=0)
    >>> optsys.zMakeIdealThinLens(ln, stop_shift=5)
    >>> optsys.zMakeIdealThinLens(ln, stop_pos=1, stop_shift=5)

    Notes
    -----
    1. For more information see "Paraxial" under "Sequential surface type
       definitions" in the Zemax manual.
    2. Use ``ln.zPushLens(1)`` update lens into the LDE
    """
    epd = fl/fn
    ddeLn.zNewLens()
    ddeLn.zSetSystemAper(aType=0, stopSurf=1, apertureValue=epd)
    ddeLn.zSetSystemProperty(code=18, value1=1) # Afocal Image Space
    ddeLn.zSetSurfaceData(surfaceNumber=0, code=3, value=fl)
    ddeLn.zSetSurfaceData(surfaceNumber=1, code=0, value='PARAXIAL')
    ddeLn.zSetSurfaceData(surfaceNumber=1, code=3, value=ima_dist)
    ddeLn.zSetSurfaceParameter(surfaceNumber=1, parameter=1, value=fl) # focallength
    ddeLn.zSetSurfaceParameter(surfaceNumber=1, parameter=2, value=opd_mode)
