# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        systems.py
# Purpose:     Simple optical systems for quick setup with PyZDDE.
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
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
    >>> import pyzdde.systems as optsys
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
    ddeLn.zSetSystemAper(aType=0, stopSurf=stop_surf, aperVal=epd)
    if stop_pos:
        ddeLn.zSetSurfaceData(para_surf, code=3, value=stop_shift)
    else:
        ddeLn.zSetSurfaceData(stop_surf, code=3, value=stop_shift)
    ddeLn.zSetSurfaceData(para_surf, code=0, value='PARAXIAL')
    ddeLn.zSetSurfaceParameter(para_surf, param=1, value=fl) # focallength
    ddeLn.zSetSurfaceParameter(para_surf, param=2, value=opd_mode)
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
    ddeLn.zSetSystemAper(aType=0, stopSurf=1, aperVal=epd)
    ddeLn.zSetSystemProperty(code=18, value1=1) # Afocal Image Space
    ddeLn.zSetSurfaceData(surfNum=0, code=3, value=fl)
    ddeLn.zSetSurfaceData(surfNum=1, code=0, value='PARAXIAL')
    ddeLn.zSetSurfaceData(surfNum=1, code=3, value=ima_dist)
    ddeLn.zSetSurfaceParameter(surfNum=1, param=1, value=fl) # focallength
    ddeLn.zSetSurfaceParameter(surfNum=1, param=2, value=opd_mode)

def zMakeBeamExpander(ddeLn, inDia=5.0, outDia=10.0, expGlass='N-BK7', expThick=10.0,
                      colGlass='N-BK7', colThick=10.0, preExpThick=5.0, preColThick=200.0,
                      preImgThick=10.0, insertAfter=None, insertOperandRow=1, epd=None,
                      setSysAper=True, afocal=True):
    """Creates a basic beam expander system consisting of a expander and
    a collimator lens (totally 5 surfaces excluding OBJ and IMA)

    Parameters
    ----------
    inDia : float, optional
        input beam diameter in lens units (Default = 5.0 mm)
    outDia : float
        output beam diameter in lens units  (Default = 10.0 mm)
    expGlass : string
        expander glass type (Default = 'N-BK7')
    expThick : float, optional
        thickness of the expander surface (Default = 10.0 mm)
    colGlass : string, optional
        collimator glass type (Default = 'N-BK7')
    preExpThick : float, optional
        thickness of surface before the first collimator surface, in mm.
        (Default = 5.0 mm)
    preColThick : float, optional
        thickness of the surface before collimator lens (between expander
        and collimator, Default = 200.0 mm)
    preImgThick : float, optional
        thickness of the surface after collimator (Default = 10.0 mm)
    insertAfter : integer, optional
        surface number after which to insert beam expander system
    insertOperandRow : integer, optional
        row number in MFE to insert the "REAY" operand for output beam
    epd : float, optional
        entrance pupil diameter, if desirable to set a value different
        from ``inDia``. If not provided, the ``EPD`` value of the system
        is set to ``inDia`` value.
    setSysAper : bool, optional
        by default the system aperture is set in the fucntion, such that
        the system aperture is "Entrance pupil diameter", and the value
        of the EPD is dependent on ``inDia`` and ``epd``.
    afocal : bool, optional
        by default the system is set to be image space afocal system. If
        False, then this setting is skipped.

    Returns
    -------
    None

    Examples
    --------
    >>> import pyzdde.zdde as pyz
    >>> ln = pyz.createLink()
    >>> optsys.zMakeBeamExpander(ln)

    Notes
    -----
    The system created by this function is not optimized. It only sets up
    the system and places variable solves on the radii of the appropriate
    surfaces. Please set up the MFE as described in [HTA]_ and optimize.

    References
    ----------
    The beam expander created by this function is similar to the one shown
    in the Zemax knowledgebase article [HTA]_


    .. [HTA] How to design Afocal Systems. Link: http://kb-en.radiantzemax.com/Knowledgebase/How-to-Design-Afocal-Systems
    """
    epd = inDia if epd is None else epd
    # the stop surface is the first surface of the Expander
    stop_surf = ddeLn.zGetSystemAper().stopSurf
    if insertAfter is None:
        ddeLn.zNewLens()
    if setSysAper:
        ddeLn.zSetSystemAper(aType=0, stopSurf=stop_surf, aperVal=epd)
    if afocal:
        ddeLn.zSetSystemProperty(code=18, value1=1) # Afocal Image Space
    # instert surfaces
    if insertAfter:
        in_beam_surf = insertAfter + 1
        ddeLn.zInsertSurface(in_beam_surf)
    else:
        in_beam_surf = stop_surf
    ddeLn.zSetSurfaceData(in_beam_surf, code=1, value="input beam")
    ddeLn.zSetSurfaceData(in_beam_surf, code=3, value=preExpThick)
    exp_ft_surf, exp_bk_surf = in_beam_surf + 1, in_beam_surf + 2
    col_ft_surf, col_bk_surf = in_beam_surf + 3, in_beam_surf + 4
    ddeLn.zInsertSurface(exp_ft_surf)
    ddeLn.zSetSurfaceData(exp_ft_surf, code=1, value="expander")
    ddeLn.zSetSurfaceData(exp_ft_surf, code=3, value=expThick)
    ddeLn.zSetSurfaceData(exp_ft_surf, code=4, value=expGlass)
    ddeLn.zInsertSurface(exp_bk_surf)
    ddeLn.zSetSurfaceData(exp_bk_surf, code=1, value=" ")
    ddeLn.zSetSurfaceData(exp_bk_surf, code=3, value=preColThick)
    ddeLn.zInsertSurface(col_ft_surf)
    ddeLn.zSetSurfaceData(col_ft_surf, code=1, value="collimator")
    ddeLn.zSetSurfaceData(col_ft_surf, code=3, value=colThick)
    ddeLn.zSetSurfaceData(col_ft_surf, code=4, value=colGlass)
    ddeLn.zInsertSurface(col_bk_surf)
    ddeLn.zSetSurfaceData(col_bk_surf, code=1, value=" ")
    ddeLn.zSetSurfaceData(col_bk_surf, code=3, value=preImgThick)
    # Set variable solves on the Radii of the
    for surf in [exp_ft_surf, exp_bk_surf, col_ft_surf, col_bk_surf]:
        ddeLn.zSetSolve(surf, 0, 1)
    # Set MFE operand on for the output diameter of the beam
    pWave = ddeLn.zGetPrimaryWave()
    ddeLn.zInsertMFO(insertOperandRow)
    ddeLn.zSetOperandRow(insertOperandRow, 'REAY', int1=col_bk_surf+1,
                         int2=pWave, data4=1.0, tgt=outDia/2.0, wgt=1.0)
