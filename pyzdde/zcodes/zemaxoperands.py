#-------------------------------------------------------------------------------
# Name:        zemaxoperands.py
# Purpose:     Class of ZEMAX operands
#
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2012 - 2014
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
# Revision:    0.7.2
#-------------------------------------------------------------------------------
from __future__ import print_function
import re as _re
from pyzdde.utils.pyzddeutils import _prettifyCodeDesc, _boldifyText, _prettifyText, _print_mod

class _Operands(object):
    """
    An operand is a single command which ZEMAX uses to perform a calculation or
    function. There are three types of operands in ZEMAX: optimization,
    tolerancing, and multi-configuration. Each operand is a four character code.
    For example, the optimization operand TRAR stands for TRansverse Aberration
    Radius. The TRAR operand, when listed in the Merit Function Editor, causes a
    single ray to be traced whose aberration value is returned in the "value"
    column.
    The list of operands were compiled from ZEMAX Version 13.0404, 2013.
    """
    opt_operands = {  # key is operand type, value is a short description
    "ABCD": ("The ABCD values used by the grid distortion feature to compute "
            "generalized distortion. See also DISA."),
    "ABGT":"Absolute value of operand greater than.",
    "ABLT": "Absolute value of operand less than.",
    "ABSO": "Absolute value of the operand defined by Op#.",
    "ACOS": "Arccosine of the value of the operand defined by Op#.",
    "AMAG": "Angular magnification.",
    "ANAC": ("Angular aberration radial direction measured in image space with "
            "respect to the centroid at the wavelength defined by Wave."),
    "ANAR": ("Angular aberration radius measured in image space at the "
            "wavelength defined by Wave with respect to the primary wavelength "
            "chief ray."),
    "ANAX": ("Angular aberration x direction measured in image space at the "
            "wavelength defined by Wave with respect to the primary wavelength "
            "chief ray."),
    "ANAY": ("Angular aberration y direction measured in image space at the "
            "wavelength defined by Wave with respect to the primary wavelength "
            "chief ray."),
    "ANCX": ("Angular aberration x direction measured in image space at the "
            "wavelength defined by Wave with respect to the centroid."),
    "ANCY": ("Angular aberration y direction measured in image space at the "
            "wavelength defined by Wave with respect to the centroid."),
    "ASIN": "Arcsine of the value of the operand defined by Op#.",
    "ASTI": ("Astigmatism in waves contributed by the surface defined by Surf at "
            "the wavelength defined by Wave."),
    "ATAN": "Arctangent of the value of the operand defined by Op#.",
    "AXCL": ("Axial color, measured in lens units for focal systems and diopters "
            "for afocal systems."),
    "BFSD": "Best Fit Sphere (BFS) data.",
    "BIOC": ("Biocular Convergence. Returns the convergence between two eye "
            "configurations in milliradians."),
    "BIOD": ("Biocular Dipvergence. Returns the dipvergence between two eye "
            "configurations in milliradians."),
    "BIPF": "Unused.",
    "BLNK": "Does nothing. Used for separating portions of the operand list.",
    "BSER": "Boresight error.",
    "CEGT": ("Boundary operand that constrains the coating extinction offset to "
            "be greater than the target value."),
    "CELT": ("Boundary operand that constrains the coating extinction offset to "
            "be less than the target value."),
    "CENX": "Centroid X position. See also CENY, CNPX, CNPY, CNAX, and CNAY.",
    "CENY": "Centroid Y position. See also CENX, CNPX, CNPY, CNAX, and CNAY.",
    "CEVA": ("Boundary operand that constrains the coating extinction offset to "
            "be equal to the target value."),
    "CIGT": ("Boundary operand that constrains the coating index offset of the "
            "coating layer to be greater than the target value."),
    "CILT": ("Boundary operand that constrains the coating index offset of the "
            "coating layer to be less than the target value."),
    "CIVA": ("Boundary operand that constrains the coating index offset of the "
            "coating layer to be equal to the target value."),
    "CMFV": "Construction merit function value.",
    "CMGT": ("Boundary operand that constrains the coating multiplier of the "
            "coating layer to be greater than the target value."),
    "CMLT": ("Boundary operand that constrains the coating multiplier of the "
            "coating layer to be less than the target value."),
    "CMVA": ("Boundary operand that constrains the coating multiplier of the "
            "coating layer to be equal to the target value."),
    "CNAX": "Centroid angular x direction. See also CNAY, CNPX, CNPY, CENX, CENY.",
    "CNAY": "Centroid angular y direction. See CNAX.",
    "CNPX": "Similar to CNAX, but computes the centroid position rather than angle.",
    "CNPY": "Similar to CNAY, but computes the centroid position rather than angle.",
    "CODA": "Coating Data.",
    "COGT": ("Boundary operand that constrains the conic of the surface to be "
             "greater than the specified target value."),
    "COLT": ("Boundary operand that constrains the conic of the surface to be "
            "less than the specified target value."),
    "COMA": ("Coma in waves contributed by the surface defined by Surf at the "
            "wavelength defined by Wave."),
    "CONF": "Configuration.",
    "CONS": "Constant value.",
    "COSA": "Unused.",
    "COSI": "Cosine of the value of the operand defined by Op#.",
    "COVA": "Conic value. Returns the conic constant of the surface defined by Surf.",
    "CTGT": "Center thickness greater than. See also MNCT.",
    "CTLT": "Center thickness less than. See also MXCT.",
    "CTVA": "Center thickness value.",
    "CVGT": "Curvature greater than.",
    "CVIG": "Clears the vignetting factors.See also `SVIG`.",
    "CVLT": "Curvature less than.",
    "CVOL": "Cylinder volume.",
    "CVVA": "Curvature value.",
    "DENC": "Diffraction  Encircled  Energy  (distance). See also DENF, GENC and XENC.",
    "DENF": "Diffraction Encircled Energy (fraction). See also DENC, GENC, GENF, and XENC.",
    "DIFF": "Difference of two operands (Op#1 - Op#2).",
    "DIMX": "Distortion maximum. See also DIST,",
    "DISA": "Distortion, ABCD. See also `ABCD` and `DISG`.",
    "DISC": "Distortion, calibrated.",
    "DISG": "Generalized distortion, either in percent or as an absolute distance.",
    "DIST": ("Distortion in waves contributed by the surface defined by Surf at "
            "the wavelength defined by Wave. See also DISG."),
    "DIVB": ("Divides the value of any prior operand defined by Op# by any factor"
            " defined by Factor."),
    "DIVI": "Division of first by second operand (Op#1 / Op#2). See also 'RECI'.",
    "DLTN": ("Delta N. Computes the diff. between the max & min index of "
            "refraction on axis for a gradient index surface."),
    "DMFS": "Default merit function start.",
    "DMGT": "Diameter greater than.",
    "DMLT": "Diameter less than.",
    "DMVA": "Diameter value.",
    "DXDX": "Derivative of transverse x-aberration with respect to x-pupil coordinate.",
    "DXDY": "Derivative of transverse x-aberration with respect to y-pupil coordinate.",
    "DYDX": "Derivative of transverse y-aberration with respect to x-pupil coordinate.",
    "DYDY": "Derivative of transverse y-aberration with respect to y-pupil coordinate.",
    "EFFL": "Effective focal length in lens units. The wavelength used is defined by Wave.",
    "EFLX": ("Effective focal length in the local x plane of the range of "
            "surfaces defined by Surf1 and Surf2 at the primary wavelength."),
    "EFLY": ("Effective focal length in the local y plane of the range of "
            "surfaces defined by Surf1 and Surf2 at the primary wavelength."),
    "EFNO": "Effective F/#. See also RELI.",
    "ENDX": "End execution.",
    "ENPP": "Entrance pupil position in lens units, with respect to the first surface.",
    "EPDI": "Entrance pupil diameter in lens units.",
    "ERFP": "Edge Response Function Position.",
    "EQUA": "Equal operand. See SUMM and OSUM.",
    "ETGT": "Edge thickness greater than. See also MNET.",
    "ETLT": "Edge thickness less than. See also MXET.",
    "ETVA": "Edge thickness value.See also MNET.",
    "EXPD": "Exit pupil diameter in lens units.",
    "EXPP": "Exit pupil position in lens units, with respect to the image surface.",
    "FCGS": "Generalized field curvature, sagittal. See also FCGT",
    "FCGT": "Generalized field curvature, tangential. See also FCGS.",
    "FCUR": "Field curvature in waves contributed by the surface.",
    "FDMO": "Field Data Modify.",
    "FDRE": "Field Data Restore. See FDMO.",
    "FICL": "Fiber coupling efficiency for single mode fibers. See also FICP.",
    "FICP": "Fiber coupling as computed using the (POP) algorithm. See also FICL.",
    "FOUC": "Foucault analysis.",
    "FREZ": "Freeform Z object constraints.",
    "FTGT": "Full thickness greater than. See FTLT.",
    "FTLT": "Full thickness less than. See FTGT.",
    "GBPD": "Gaussian beam (paraxial) divergence in the optical space following "
            "the surface.",
    "GBPP": "Gaussian beam (paraxial) position, which is the distance from the "
            "waist to the surface. See GBPD.",
    "GBPR": "Gaussian beam (paraxial) radius of curvature in the optical space "
            "following the specified surface.See GBPD.",
    "GBPS": "Gaussian beam (paraxial) size in the optical space following the "
            "specified surface. See GBPD.",
    "GBPW": "Gaussian beam (paraxial) waist in the optical space following the "
            "specified surface. See GBPD.",
    "GBSD": "Gaussian beam (skew) divergence in the optical space following the "
            "specified surface.",
    "GBSP": "Gaussian beam (skew) position, which is the distance from the waist "
            "to the surface. See GBSD.",
    "GBSR": "Gaussian beam (skew) radius in the optical space following the "
            "specified surface. See GBSD.",
    "GBSS": "Gaussian beam (skew) size in the optical space following the "
            "specified surface. See GBSD.",
    "GBSW": "Gaussian beam (skew) waist in the optical space following the "
            "specified surface. See GBSD.",
    "GCOS": "Glass cost.",
    "GENC": "Geometric Encircled Energy (distance). See also GENF, DENC, DENF, "
            "and XENC.",
    "GENF": "Geometric Encircled Energy (fraction). See also GENC, DENC, DENF, "
            "and XENC.",
    "GLCA": "Global x-direction orientation vector component of the surface "
            "defined by Surf.",
    "GLCB": "Global y-direction orientation vector component of the surface "
            "defined by Surf.",
    "GLCC": "Global z-direction orientation vector component of the surface "
            "defined by Surf.",
    "GLCR": "Global Coordinate Rotation Matrix component at the surface "
            "defined by Surf.",
    "GLCX": "Global vertex x-coordinate of the surface defined by Surf.",
    "GLCY": "Global vertex y-coordinate of the surface defined by Surf.",
    "GLCZ": "Global vertex z-coordinate of the surface defined by Surf.",
    "GMTA": "Geometric MTF average of sagittal and tangential response.",
    "GMTS": "Geometric MTF sagittal response. See GMTA.",
    "GMTT": "Geometric MTF tangential response. See GMTA.",
    "GOTO": "Skips all operands between the GOTO operand line and the operand "
            "number defined by Op#.",
    "GPIM": "Ghost pupil image. See also GPRT, GPRX, GPRY, GPSX, and GPSY.",
    "GPRT": "Ghost ray transmission.",
    "GPRX": "Ghost ray real x coordinate.",
    "GPRY": "Ghost ray real y coordinate.",
    "GPSX": "Ghost ray paraxial x coordinate.",
    "GPSY": "Ghost ray paraxial y coordinate.",
    "GRMN": "Gradient index minimum index.",
    "GRMX": "Gradient index maximum index.",
    "GTCE": "Glass TCE. For non-glass surfaces, see `TCVA`.",
    "HACG": "Unused.",
    "HHCN": "Test for the hyperhemisphere condition.",
    "IMAE": "Image analysis data.",
    "IMSF": "Image  Surface",
    "INDX": "Index of refraction.",
    "InGT": "Index `n` greater than.",
    "InLT": "Index `n` less than.",
    "InVA": "This operand is similar to InGT except it constrains the current "
            "value of the index of refraction.",
    "ISFN": "Image space F/#. See WFNO.",
    "ISNA": "Image space numerical aperture. See ISFN.",
    "LACL": "Lateral color.",
    "LINV": "Lagrange (or optical) invariant of system in lens units at the "
            "wavelength defined by Wave.",
    "LOGE": "Log base e of an operand.",
    "LOGT": "Log base 10 of an operand.",
    "LONA": "Longitudinal aberration, measured in lens units for focal systems "
            "and diopters for afocal systems. See AXCL.",
    "LPTD": "This boundary operand constrains the slope of the axial grad. ind. "
            "profile from changing signs within a gradient index component.",
    "MAXX": "Returns the largest value within the indicated range of operands "
            "defined by Op#1 and Op#2 . See MINN.",
    "MCOG": "Multi-configuration operand greater than. See MCOL.",
    "MCOL": "Multi-configuration operand less than. See MCOG.",
    "MCOV": "Multi-configuration operand value. See MCOG.",
    "MINN": "Returns the smallest value within the indicated range of operands. "
            "See MAXX.",
    "MNAB": "Minimum Abbe number. See also MXAB.",
    "MNCA": "Minimum center thickness air. See also  MNCT  and  MNCG.",
    "MNCG": "Minimum center thickness glass. See also MNCT and MNCA.",
    "MNCT": "Minimum center thickness.",
    "MNCV": "Minimum curvature.",
    "MNDT": "Minimum diameter to thickness ratio. See also MXDT.",
    "MNEA": "Minimum edge thickness air. See also MNET, MNEG, ETGT, and XNEA.",
    "MNEG": "Minimum edge thickness glass. See also MNET, MNEA, ETGT, and XNEG.",
    "MNET": "Minimum edge thickness. See also MNEG, MNEA, ETGT, and XNET.",
    "MNIN": "Minimum index at d-light.",
    "MNPD": "Minimum.",
    "MNSD": "Minimum semi-diameter.",
    "MSWA": "Modulation square-wave transfer function, average of sagittal and "
            "tangential. See MTFA for details.",
    "MSWS": "Modulation square-wave transfer function, sagittal. See MTFA for "
            "details.",
    "MSWT": "Modulation square-wave transfer function, tangential. See MTFA for "
            "details.",
    "MTFA": "Diffraction modulation transfer function, average of sagittal and "
            "tangential.",
    "MTFS": "Modulation transfer function, sagittal. See MTFA for details.",
    "MTFT": "Modulation transfer function, tangential. See MTFA for details.",
    "MTHA": "Huygens Modulation transfer function, average of sagittal and "
            "tangential.",
    "MTHS": "Huygens Modulation transfer function, sagittal. See MTHA for details.",
    "MTHT": "Huygens Modulation transfer function, tangential. See MTHA for details.",
    "MXAB": "Maximum Abbe number.",
    "MXCA": "Maximum  center  thickness air.",
    "MXCG": "Maximum center thickness glass.",
    "MXCT": "Maximum center thickness.",
    "MXCV": "Maximum curvature. See also MNCV.",
    "MXDT": "Maximum diameter to thickness ratio. See also MNDT.",
    "MXEA": "Maximum edge thickness air. See also MXET, MXEG, ETLT, and XXEA.",
    "MXEG": "Maximum edge thickness glass.See also MXET, MXEA, ETLT, and XXEG.",
    "MXET": "Maximum  edge  thickness. See also `MXEG`, `MXEA`,`ETLT`, and `XXET`.",
    "MXIN": "Maximum index at d-light. See also MNIN.",
    "MXPD": "Maximum. See also MNPD.",
    "MXSD": "Maximum semi-diameter.",
    "NORD": "Normal distance to the next surface.",
    "NORX": "Normal vector x component.",
    "NORY": "Normal vector y component.",
    "NORZ": "Normal vector z component.",
    "NPGT": "Non-sequential parameter greater than.",
    "NPLT": "Non-sequential parameter less than. See NPGT.",
    "NPVA": "Non-sequential parameter value. See NPGT.",
    "NPXG": "Non-sequential object position x greater than.",
    "NPXL": "Non-sequential object position x less than. See NPXG.",
    "NPXV": "Non-sequential object position x value. See NPXG.",
    "NPYG": "Non-sequential object position y greater than. See NPXG.",
    "NPYL": "Non-sequential object position y less than. See NPXG.",
    "NPYV": "Non-sequential object position y value. See NPXG.",
    "NPZG": "Non-sequential object position z greater than. See NPXG.",
    "NPZL": "Non-sequential object position z less than. See NPXG.",
    "NPZV": "Non-sequential object position z value. See NPXG.",
    "NSDC": "Non-sequential coherent data.",
    "NSDD": "Non-sequential incoherent intensity data.",
    "NSDE": "Non-sequential Detector Color object data. See also NSDD and NSDP.",
    "NSDP": "Non-sequential Detector Polar object data. See also NSDD and NSDE.",
    "NSRA": "Non-sequential single ray trace.",
    "NSRM": "Non-sequential Rotation Matrix component.",
    "NSST": "Non-sequential single ray trace. See also NSTR.",
    "NSTR": "Non-sequential trace. See also NSST.",
    "NTXG": "Non-sequential object tilt about x greater than. See NPXG.",
    "NTXL": "Non-sequential object tilt about x less than. See NPXG.",
    "NTXV": "Non-sequential object tilt about x value. See NPXG.",
    "NTYG": "Non-sequential object tilt about y greater than. See NPXG.",
    "NTYL": "Non-sequential object tilt about y less than. See NPXG.",
    "NTYV": "Non-sequential object tilt about y value. See NPXG.",
    "NTZG": "Non-sequential object tilt about z greater than. See NPXG.",
    "NTZL": "Non-sequential object tilt about z less than. See NPXG.",
    "NTZV": "Non-sequential object tilt about z value. See NPXG.",
    "OBSN": "Object space numerical aperture.",
    "OOFF": "This operand indicates an unused entry in the operand list.",
    "OPDC": "Optical path difference with respect to chief ray in waves at the "
            "wavelength defined by Wave.",
    "OPDM": "Optical path difference with respect to the mean OPD over the pupil "
            "at the wavelength defined by Wave.",
    "OPDX": "Optical path difference with respect to the mean OPD over the pupil "
            "with tilt removed at the wavelength defined by Wave.",
    "OPGT": "Operand greater than.",
    "OPLT": "Operand less than.",
    "OPTH": "Optical path length. See PLEN.",
    "OPVA": "Operand value.",
    "OSCD": "Offense against the sine condition (OSC) at the wavelength defined "
            "by Wave.",
    "OSUM": "Sums the values of all operands between the two operands defined by "
            "Op#1 and Op#2. See SUMM.",
    "PANA": "Paraxial ray x-direction surface normal at the ray-surface intercept "
            "at the wavelength defined by Wave.",
    "PANB": "Paraxial ray y-direction surface normal at the ray-surface intercept "
            "at the wavelength defined by Wave.",
    "PANC": "Paraxial ray z-direction surface normal at the ray-surface intercept "
            "at the wavelength defined by Wave.",
    "PARA": "Paraxial ray x-direction cosine of the ray after refraction from "
            "the surface defined by Surf at the wavelength defined by Wave.",
    "PARB": "Paraxial ray y-direction cosine of the ray after refraction from "
            "the surface defined by Surf at the wavelength defined by Wave.",
    "PARC": "Paraxial ray z-direction cosine of the ray after refraction from "
            "the surface defined by Surf at the wavelength defined by Wave.",
    "PARR": "Paraxial ray radial coordinate in lens units at the surface defined "
            "by Surf at the wavelength defined by Wave.",
    "PARX": "Paraxial ray x-coordinate in lens units at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "PARY": "Paraxial ray y-coordinate in lens units at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "PARZ": "Paraxial ray z-coordinate in lens units at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "PATX": "Paraxial ray x-direction ray tangent.",
    "PATY": "Paraxial ray y-direction ray tangent.",
    "PETC": "Petzval curvature in inverse lens units at the wavelength defined "
            "by Wave.",
    "PETZ": "Petzval radius of curvature in lens units at the wavelength defined "
            "by Wave.",
    "PIMH": "Paraxial image height at the paraxial image surface at the "
            "wavelength defined by Wave.",
    "PLEN": "Path length. See OPTH.",
    "PMAG": "Paraxial magnification.",
    "PMGT": "Parameter greater than.",
    "PMLT": "Parameter less than.",
    "PMVA": "Parameter value.",
    "PnGT": "This operand is obsolete, use PMGT instead.",
    "PnLT": "This operand is obsolete, use PMLT instead.",
    "PnVA": "This operand is obsolete, use PMVA instead.",
    "POPD": "Physical Optics Propagation Data.",
    "POPI": "Physical Optics Propagation Data.",
    "POWF": "Power at a field point.",
    "POWP": "Power at a point in the pupil.",
    "POWR": "The surface power (in inverse lens units) of the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "PRIM": "Primary wavelength. This is used to change the primary wavelength "
            "number to the wavelength defined by Wave during merit function "
            "evaluation.",
    "PROB": "Multiplies the value of the operand defined by Op# by the factor "
            "defined by Factor.",
    "PROD": "Product of two operands (Op#1 X Op#2). See PROB.",
    "QOAC": "Unused.",
    "QSUM": "Quadratic sum. See also SUMM, OSUM, EQUA.",
    "RAED": "Real ray angle of exitance. See also RAID.",
    "RAEN": "Real ray angle of exitance. See also RAIN.",
    "RAGA": "Global ray x-direction cosine.",
    "RAGB": "Global ray y-direction cosine. See RAGA.",
    "RAGC": "Global ray z-direction cosine. See RAGA.",
    "RAGX": "Global ray x-coordinate.",
    "RAGY": "Global ray y-coordinate. See RAGX.",
    "RAGZ": "Global ray z-coordinate. See RAGX.",
    "RAID": "Real ray angle of incidence.See also RAED.",
    "RAIN": "Real ray angle of incidence.See also RAEN.",
    "RANG": "Ray angle in radians with respect to z axis.",
    "REAA": "Real ray x-direction cosine of the ray after refraction from the "
            "surface defined by Surf at the wavelength defined by Wave.",
    "REAB": "Real ray y-direction cosine of the ray after refraction from the "
            "surface defined by Surf at the wavelength defined by Wave.",
    "REAC": "Real ray z-direction cosine of the ray after refraction from the "
            "surface defined by Surf at the wavelength defined by Wave.",
    "REAR": "Real ray radial coordinate in lens units at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "REAX": "Real ray x-coordinate in lens units at the surface defined by Surf "
            "at the wavelength defined by Wave.",
    "REAY": "Real ray y-coordinate in lens units at the surface defined by Surf "
            "at the wavelength defined by Wave.",
    "REAZ": "Real ray z-coordinate in lens units at the surface defined by Surf "
            "at the wavelength defined by Wave.",
    "RECI": "Returns the reciprocal of the value of operand Op#1. See also `DIVI`.",
    "RELI": "Relative illumination. See also EFNO.",
    "RENA": "Real ray x-direction surface normal at the ray-surface intercept "
            "at the surfaced defined by Surf at the wavelength defined by Wave.",
    "RENB": "Real ray y-direction surface normal at the ray-surface intercept "
            "at the surface defined by Surf at the wavelength defined by Wave.",
    "RENC": "Real ray z-direction surface normal at the ray-surface intercept "
            "at the surface defined by Surf at the wavelength defined by Wave.",
    "RETX": "Real ray x-direction ray tangent (slope) at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "RETY": "Real ray y-direction ray tangent (slope) at the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "RGLA": "Reasonable glass.",
    "RSCE": "RMS spot radius with respect to the centroid in lens units.",
    "RSCH": "RMS spot radius with respect to the chief ray in lens units.",
    "RSRE": "RMS spot radius with respect to the centroid in lens units.",
    "RSRH": "RMS spot radius with respect to the chief ray in lens units.",
    "RWCE": "RMS wavefront error with respect to the centroid in waves.",
    "RWCH": "RMS wavefront error with respect to the chief ray in waves.",
    "RWRE": "RMS wavefront error with respect to the centroid in waves.",
    "RWRH": "RMS wavefront error with respect to the chief ray in waves.",
    "SAGX": "The sag in lens units of the surface defined by Surf at X = the "
            "semi-diameter, and Y = 0. See also SSAG.",
    "SAGY": "The sag in lens units of the surface defined by Surf at Y = the "
            "semi-diameter, and X = 0. See also SSAG.",
    "SCUR": "Surface curvature.",
    "SFNO": "Sagittal working F/#, computed at the field point defined by Field "
            "and the wavelength defined by Wave. See TFNO.",
    "SINE": "Sine of the value of the operand defined by Op#. If Flag is 0, then "
            "the units are radians, otherwise, degrees.",
    "SKIN": "Skip if not symmetric. See SKIS.",
    "SKIS": "Skip if symmetric. If the lens is rotationally symmetric, then "
            "computation of the merit function continues at the operand defined by Op#.",
    "SMIA": "SMIA-TV Distortion.",
    "SPCH": "Spherochromatism in lens units.",
    "SPHA": "Spherical aberration in waves contributed by the surface defined by "
            "Surf at the wavelength defined by Wave.",
    "SQRT": "Square root of the operand defined by Op#.",
    "SSAG": "The sag in lens units of the surface defined by Surf at the coordinate "
            "defined by X and Y. See also SAGX, SAGY.",
    "STHI": "Surface Thickness.",
    "STRH": "Strehl Ratio.",
    "SUMM": "Sum of two operands (Op#1 + Op#2). See OSUM.",
    "SVIG": "Sets the vignetting factors for the current configuration. "
            "See also `CVIG`.",
    "TANG": "Tangent of the value of the operand defined by Op#.",
    "TCGT": "Thermal Coefficient of expansion greater than.",
    "TCLT": "Thermal Coefficient of expansion less than.",
    "TCVA": "Thermal Coefficient of expansion value. For glass surfaces, see `GTCE`.",
    "TFNO": "Tangential working F/#, computed at the field point defined by "
            "Field and the wavelength defined by Wave. See SFNO.",
    "TGTH": "Sum of glass thicknesses from Surf1 to Surf2. See TTHI.",
    "TMAS": "Total mass.",
    "TOLR": "Tolerance data.",
    "TOTR": "Total track (length) of lens in lens units. See `Total track`.",
    "TRAC": "Transverse aberration radial direction measured in image space with "
            "respect to the centroid for the wavelength  defined  by  Wave.",
    "TRAD": "The x component of the TRAR only. TRAD has the same restrictions that "
            "TRAC does; see TRAC for a detailed discussion.",
    "TRAE": "The y component of the TRAR only. TRAE has the same restrictions that "
            "TRAC does; see TRAC for a detailed discussion.",
    "TRAI": "Transverse aberration radius measured at the surface defined by Surf "
            "at the wavelength defined by Wave with respect to the chief ray.",
    "TRAN": "Unused.",
    "TRAR": "Transverse aberration radial direction measured in image space at the "
            "wavelength defined by Wave with respect to the chief ray. See ANAR.",
    "TRAX": "Transverse aberration x direction measured in image space at the "
            "wavelength defined by Wave with respect to the chief ray.",
    "TRAY": "Transverse aberration y direction measured in image space at the "
            "wavelength defined by Wave with respect to the chief ray.",
    "TRCX": "Transverse aberration x direction measured in image space with "
            "respect to the centroid. TRCX has the same restrictions that TRAC does.",
    "TRCY": "Transverse aberration y direction measured in image space with "
            "respect to the centroid. TRCY has the same restrictions that TRAC does.",
    "TTGT": "Total  thickness  greater  than. See TTLT and TTVA.",
    "TTHI": "Sum of thicknesses of surfaces from Surf1 to Surf2. See TGTH.",
    "TTLT": "Total thickness less than. See TTGT.",
    "TTVA": "Total thickness value. See TTGT.",
    "UDOP": ("User defined operand. Used for optimizing numerical results "
            "computed in externally compiled programs. See also ZPLM."),
    "USYM": ("If present in merit function, instructs ZEMAX to assume radial "
            "symmetry exists in the lens even if ZEMAX detects symmetry does not."),
    "VOLU": "Volume of element(s) in cubic cm.",
    "WFNO": "Working F/#. See `Working F/#` on page 62, and ISFN, SFNO, and TFNO.",
    "WLEN": "Wavelength. This operand returns the wavelength defined by Wave "
            "in micrometers.",
    "XDGT": "Extra data value greater than.",
    "XDLT": "Extra data value less than.",
    "XDVA": "Extra data value.",
    "XENC": "Extended source encircled energy (distance). See also XENF, DENC, "
            "DENF, GENC, and GENF.",
    "XENF": "Extended source encircled energy (fraction). See also XENC, GENC, "
            "GENF, DENC, and DENF.",
    "XNEA": "Minimum edge thickness for the range of air surfaces defined by "
            "Surf1 and Surf2. See MNEA.",
    "XNEG": "Minimum edge thickness for the range of glass surfaces defined "
            "by Surf1 and Surf2.",
    "XNET": "Minimum edge thickness for the range of surfaces defined by Surf1 "
            "and Surf2. See MNET.",
    "XXEA": "Maximum edge thickness for the range of air surfaces defined by "
            "Surf1 and Surf2. See MXEA.",
    "XXEG": "Maximum edge thickness for the range of glass surfaces defined by "
            "Surf1 and Surf2. See MXEG.",
    "XXET": "Maximum  edge thickness for the range of surfaces defined by  "
            "Surf1  and Surf2. See MXET.",
    "YNIP": "YNI-paraxial. It is the product of the parax. marg. ray ht. & the "
            "index times the angle of incidence at the surface.",
    "ZERN": "Zernike Fringe coefficient.",
    "ZPLM": "Used for optimizing numerical results computed in ZPL macros. "
            "See also UDOP.",
    "ZTHI": "This operand controls the variation in the total thickness of the "
            "range surfaces defined by Surf1 and Surf2 over multiple configurations.",
    }

    tol_operands = { # key is operand type, value is a short description
    "TRAD": "Tolerance on surface radius of curvature in lens units",
    "TCUR": "Tolerance on surface curvature in inverse lens units",
    "TFRN": "Tolerance on surface radius of curvature in fringes",
    "TTHI": "Tolerance on thickness or position in lens units",
    "TCON": "Tolerance on conic constant (dimensionless)",
    "TSDI": "Tolerance on semi-diameter in lens units",
    "TSDX": "Tolerance on Standard surface x-decenter in lens units",
    "TSDY": "Tolerance on Standard surface y-decenter in lens units",
    "TSTX": "Tolerance on Standard surface tilt in x in degrees",
    "TSTY": "Tolerance on Standard surface tilt in y in degrees",
    "TIRX": "Tolerance on Standard surface tilt in x in lens units",
    "TIRY": "Tolerance on Standard surface tilt in y in lens units",
    "TIRR": "Tolerance on Standard surface irregularity",
    "TEXI": "Tolerance on surface irregularity using Zernike Fringe polynomials",
    "TEZI": "Tolerance on surface irregularity using Zernike Standard polynomials",
    "TPAI": "Tolerance on the inverse of parameter value of surface",
    "TPAR": "Tolerance on parameter value of surface",
    "TEDV": "Tolerance on extra data value of surface",
    "TIND": "Tolerance on index of refraction of surface",
    "TABB": "Tolerance on Abbe number of surface",
    "TCMU": "Tolerance on coating multiplier of surface",
    "TCIO": "Tolerance on coating index offset of surface",
    "TCEO": "Tolerance on coating extinction offset of surface",
    "TEDX": "Tolerance on element x-decenter in lens units",
    "TEDY": "Tolerance on element y-decenter in lens units",
    "TETX": "Tolerance on element tilt about x in degrees",
    "TETY": "Tolerance on element tilt about y in degrees",
    "TETZ": "Tolerance on element tilt about z in degrees",
    "TUDX": "Tolerance on user defined x-decenter in lens units",
    "TUDY": "Tolerance on user defined y-decenter in lens units",
    "TUTX": "Tolerance on user defined tilt about x in degrees",
    "TUTY": "Tolerance on user defined tilt about y in degrees",
    "TUTZ": "Tolerance on user defined tilt about z in degrees",
    "TNPS": "Tolerance on NSC object position.",
    "TNPA": "Tolerance on an NSC object parameter data",
    "TMCO": "Tolerance on multi-configuration editor value",
    "CEDV": "Sets an extra data value as a compensator",
    "CMCO": "Sets a multi-configuration value as a compensator",
    "COMM": "This operand is used to print a comment in the tolerance analysis "
            "output report",
    "COMP": "Sets a thickness, radius or conic compensator.",
    "CPAR": "Sets a parameter as a compensator",
    "SAVE": "Saves the file used to evaluate the tolerance on the prior row in "
            "the editor",
    "SEED": "Seeds the random number generator for MC analysis. Choose 0 for random.",
    "STAT": "Defines statstics used 'on the fly' during Monte Carlo analysis",
    "TWAV": "Sets the test wavelength for operands measured in fringes",
    }

    mco_operands = {  # key is operand type, value is a short description
    "AFOC": "Afocal Image Space mode.",
    "APDF": "System apodization factor.",
    "APDT": "System apodization type.",
    "APDX": "Surface aperture X-decenter.",
    "APDY": "Surface aperture Y-decenter.",
    "APER": "System aperture value.",
    "APMN": "Surface aperture minimum value.",
    "APMX": "Surface aperture maximum value.",
    "APTP": "Surface aperture type.",
    "CONN": "Conic constant.",
    "COTN": "Coating name.",
    "CROR": "Coordinate Return Orientation.",
    "CRSR": "Coordinate Return Surface.",
    "CRVT": "Surface curvature.",
    "CSP1": "Curvature solve parameter 1.",
    "CSP2": "Curvature solve parameter 2.",
    "CWGT": "Configuration weight.",
    "EDVA": "Extra data value.",
    "FLTP": "Field type.",
    "FLWT": "Field weight.",
    "FVAN": "Field vignetting VAN factor.",
    "FVCX": "Field vignetting VCX factor.",
    "FVCY": "Field vignetting VCY factor.",
    "FVDX": "Field vignetting VDX factor.",
    "FVDY": "Field vignetting VDY factor.",
    "GCRS": "Global coordinate reference surface.",
    "GLSS": "Glass.",
    "GPJX": "Global polarization state Jx.",
    "GPJY": "Global polarization state Jy.",
    "GPIU": "Global polarization state.",
    "GPPX": "Global polarization phase Px.",
    "GPPY": "Global polarization phase Py.",
    "GQPO": "Obscuration value used for Gaussian Quadrature pupil sampling in "
            "the default merit function.",
    "HOLD": "Hold.",
    "IGNR": "Ignore This Surface status.",
    "LTTL": "Lens title.",
    "MABB": "Model glass Abbe number.",
    "MCOM": "Surface comment.",
    "MDPG": "Model glass dPgF.",
    "MIND": "Model glass index.",
    "MOFF": "Off.",
    "NCOM": "Modifies the comment for non-sequential objects in the NSC Editor.",
    "NCOT": "Non-sequential coating.",
    "NGLS": "Nonsequential object glass.",
    "NPAR": "NSCE object parameter.",
    "NPOS": "NSCE object position.",
    "NPRO": "NSC object property.",
    "PAR1": "Parameter 1. Obsolete, use PRAM instead.",
    "PAR2": "Parameter 2. Obsolete, use PRAM instead.",
    "PAR3": "Parameter 3. Obsolete, use PRAM instead.",
    "PAR4": "Parameter 4. Obsolete, use PRAM instead.",
    "PAR5": "Parameter 5. Obsolete, use PRAM instead.",
    "PAR6": "Parameter 6. Obsolete, use PRAM instead.",
    "PAR7": "Parameter 7. Obsolete, use PRAM instead.",
    "PAR8": "Parameter 8. Obsolete, use PRAM instead.",
    "PRAM": "Parameter value.",
    "PRES": "Air pressure, in atmospheres.",
    "PRWV": "Primary wavelength number.",
    "PSCX": "X Pupil Compress.",
    "PSCY": "Y Pupil Compress.",
    "PSHX": "X Pupil Shift.",
    "PSHY": "Y Pupil Shift.",
    "PSHZ": "Z Pupil Shift.",
    "PSP1": "Parameter solve parameter 1.",
    "PSP2": "Parameter solve parameter 2.",
    "PSP3": "Parameter solve parameter 3.",
    "PUCN": "Pickup range of values.",
    "RAAM": "Ray aiming.",
    "SATP": "System aperture type.",
    "SDIA": "Semi-diameter.",
    "STPS": "Stop surface number.",
    "SWCN": "Configuration number for the SolidWorks part.",
    "TCEX": "Thermal coefficient of expansion.",
    "TELE": "Telecentric in object space.",
    "TEMP": "Temperature in degrees Celsius.",
    "THIC": "Surface thickness.",
    "TSP1": "Thickness solve parameter 1.",
    "TSP2": "Thickness solve parameter 2.",
    "TSP3": "Thickness solve parameter 3.",
    "UDAF": "User-defined aperture file.",
    "WAVE": "Wavelength.",
    "WLWT": "Wavelength weight.",
    "XFIE": "X-field value.",
    "YFIE": "Y-field value.",
    }

def showZOperandList(operandType = 0):
    """Lists the operands for the specified type with a shot description of
    each operand.

    showZOperandList([operandType])->None (the operands are printed on screen)

    Parameters
    ----------
    operandType : 0 = All operands (use with caution, slow)
                  1 = Optimization operands
                  2 = Tolerancing operands (Default)
                  3 = Multi-configuration operands
    Returns
    -------
    None (the operands are printed on screen)
    """
    if operandType == 0:
        print("Listing all operands:")
        for operand, description in sorted(_Operands.opt_operands.items()):
            #print("[",operand,"]",description)
            _print_mod(_prettifyCodeDesc(operand,description))
        for elem in sorted(_Operands.tol_operands.items()):
            #print("[",operand,"]",description)
            _print_mod(_prettifyCodeDesc(operand,description))
        for elem in sorted(_Operands.mco_operands.items()):
            #print("[",operand,"]",description)
            _print_mod(_prettifyCodeDesc(operand,description))
        totOperands = (len(_Operands.opt_operands) +
                       len(_Operands.tol_operands) +
                       len(_Operands.mco_operands))
        #print("\nTotal number of operands = {:d}".format(totOperands))
        _print_mod(_boldifyText("Total number of operandss = ",str(totOperands)))
    else:
        if operandType == 1:
            print("Listing Optimization operands:")
            toList = sorted(_Operands.opt_operands.items())
        elif operandType == 2:
            print("Listing Tolerancing operands:")
            toList = sorted(_Operands.tol_operands.items())
        elif operandType == 3:
            print("Listing Multi-configuration operands:")
            toList = sorted(_Operands.mco_operands.items())
        for operand, description in toList:
            _print_mod(_prettifyCodeDesc(operand,description))
        _print_mod(_boldifyText("Total number of operands = ",str(len(toList))))

def getZOperandCount(operandType = 0):
    """Returns the total number of operands for the specified operand type

    getZOperandCount([operandType])->count

    Parameters
    ----------
    operandType : 0 = All operands
                  1 = Optimization operands
                  2 = Tolerancing operands
                  3 = Multi-configuration operands
    Returns
    -------
    count : number of operands
    """
    numOptOperands = len(_Operands.opt_operands)
    numTolOperands = len(_Operands.tol_operands)
    numMcoOperands = len(_Operands.mco_operands)
    if operandType == 0:
        return (numOptOperands + numTolOperands + numMcoOperands)
    elif operandType == 1:
        return numOptOperands
    elif operandType == 2:
        return numTolOperands
    elif operandType == 3:
        return numMcoOperands

def isZOperand(operand, operandType=0):
    """Returns True or False depending on whether the operand is a valid operand
    of the specified operendType.

    isZOperand(operand, operandType)->bool

    args:
      operand     : (string) four letter character code, such as 'ABCD'
      operandType : 0 = All operands
                    1 = Optimization operands
                    2 = Tolerancing operands
                    3 = Multi-configuration operands
    ret:
      bool : True if valid operand, else False.
    """
    if operandType == 1:
        return str(operand) in _Operands.opt_operands.viewkeys()
    elif operandType == 2:
        return str(operand) in _Operands.tol_operands.viewkeys()
    elif operandType == 3:
        return str(operand) in _Operands.mco_operands.viewkeys()
    elif operandType == 0:
        return ((str(operand) in _Operands.opt_operands.viewkeys()) or
                (str(operand) in _Operands.tol_operands.viewkeys()) or
                (str(operand) in _Operands.mco_operands.viewkeys()))
    else:
        return False

def showZOperandDescription(operand):
    """Get a short description about an operand.

    showZOperandDescription(operand)->description

    args:
      operand  : a valid ZEMAX operand
    ret:
      description : a short description of the operand and the type of operand.
    """
    if isZOperand(str(operand),1):
        _print_mod(_prettifyText(str(operand), " is a Optimization operand",
                               color0='magenta',color1='black'))
        _print_mod(_prettifyText("Description: ", _Operands.opt_operands[str(operand)],
                  color0='blue',color1='black'))
    elif isZOperand(str(operand),2):
        _print_mod(_prettifyText(str(operand), " is a Tolerancing operand",
                               color0='magenta',color1='black'))
        _print_mod(_prettifyText("Description: ", _Operands.tol_operands[str(operand)],
                  color0='blue',color1='black'))
    elif isZOperand(str(operand),3):
        _print_mod(_prettifyText(str(operand), " is a Multi-configuration operand",
                               color0='magenta',color1='black'))
        _print_mod(_prettifyText("Description: ", _Operands.mco_operands[str(operand)],
                  color0='blue',color1='black'))
    else:
        print("{} is a NOT a valid ZEMAX operand.".format(str(operand)))

def findZOperand(keywords):
    """Find/search Zemax operands using specific keywords of interest.

    findZOperand("keyword#1 [, keyword#2, keyword#3, ...]")->searchResult

    Parameters
    ----------
    keywords   : a string containing a list of comma separated keywords

    Example
    -------
    >>> zo.findZOperand("decenter")
    [TEDY] Tolerance on element y-decenter in lens units
    [TEDX] Tolerance on element x-decenter in lens units
    [TUDX] Tolerance on user defined x-decenter in lens units
    [TUDY] Tolerance on user defined y-decenter in lens units
    [TSDY] Tolerance on Standard surface y-decenter in lens units
    [TSDX] Tolerance on Standard surface x-decenter in lens units

    Found 6 Tolerance operands.

    [APDY] Surface aperture Y-decenter.
    [APDX] Surface aperture X-decenter.

    Found 2 Macro operands.

    >>> zo.findZOperand('trace')
    [NSRA] Non-sequential single ray trace.
    [NSTR] Non-sequential trace. See also NSST.
    [NSST] Non-sequential single ray trace. See also NSTR.

    Found 3 Optimization operands.
    """
    words2find = [words.strip() for words in keywords.split(",")]
    previousFoundKeys = []
    for operand, description in _Operands.opt_operands.items():
        for kWord in words2find:
            if __find(kWord,description):
                _print_mod(_prettifyCodeDesc(operand,description))
                previousFoundKeys.append(operand)
                break # break the inner for loop
    if previousFoundKeys:
        _print_mod(_boldifyText("Found ", str(len(previousFoundKeys)),
                              " Optimization operands",'blue','red','blue'))
    previousFoundKeys = []
    for operand, description in _Operands.tol_operands.items():
        for kWord in words2find:
            if __find(kWord,description):
                _print_mod(_prettifyCodeDesc(operand,description))
                previousFoundKeys.append(operand)
                break # break the inner for loop
    if previousFoundKeys:
        _print_mod(_boldifyText("Found ", str(len(previousFoundKeys)),
                              " Tolerance operands",'blue','red','blue'))
    previousFoundKeys = []
    for operand, description in _Operands.mco_operands.items():
        for kWord in words2find:
            if __find(kWord,description):
                _print_mod(_prettifyCodeDesc(operand,description))
                previousFoundKeys.append(operand)
                break # break the inner for loop
    if previousFoundKeys:
        _print_mod(_boldifyText("Found ", str(len(previousFoundKeys)),
                              " Macro operands",'blue','red','blue'))

def __find(word2find,instring):
    r = _re.compile(r'\b({0})\b'.format(word2find),flags=_re.IGNORECASE)
    if r.search(instring):
        return True
    else:
        return False

##    Possible functions to implement
##    #Get operand number, given operand type (key)
##    #Given operand number, get operand type (get key)

if __name__ == '__main__':
    #Test showZOperandList()
    showZOperandList(2)
    #Test getZOperandCount()
    print("Total number of operands:",getZOperandCount(0))
    #Test isZOperand()
    print("'TSDX' is a Tolerence oprand type (True/False):", isZOperand('TSDX',2))
    print("'TSDX' is an Oprimization oprand type (True/False):", isZOperand('TSDX',1))
    print("'RANDOM' is a Tolerence operand type (True/False):", isZOperand('RANDOM',2))
    print("'PUCN' is a Mico-configuration operand type (True/False):", isZOperand('PUCN',3))
    print("'GQPO' is a ZEMAX operand (True/False):", isZOperand('GQPO',0))
    #Test showZOperandDescription()
    showZOperandDescription('RANDOM')
    showZOperandDescription('TSDX')

