# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        zemaxbuttons.py
# Purpose:     Class of ZEMAX 3-letter Button codes
#              Last updated: April 16, 2017
# 
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
from __future__ import print_function
import re as _re
from pyzdde.utils.pyzddeutils import _prettifyCodeDesc, _boldifyText, _prettifyText, _print_mod

class _Buttons(object):
    """ZEMAX 3-letter buttons. Note ZPL Macro codes are not included.
    The list of button codes were compiled from ZEMAX Version 13.0404, 2013.
    """
    button_code = {
    "Off": "None",
    "ABg": "ABg Data Catalog",
    "Agm": "Athermal Glass Map",
    "Bac": "Backup To Archive File",
    "Bfv": "Beam File Viewer",
    "C31": "Src color chart 1931",
    "C76": "Src color chart 1976",
    "Caa": "Coating, Abs. vs. Angle",
    "Car": "Cardinal Points",
    "Cas": "Coat All Surfaces",
    "Caw": "Coating, Abs. vs. Wavelength",
    "Cca": "Convert to Circular Apertures",
    "Cda": "Coating, Dia. vs. Angle",
    "Cdw": "Coating, DIa. vs. Angle",
    "Cfa": "Convert to Floating Apertures",
    "Cfm": "Convert to Maximum Apertures",
    "Cfo": "Convert Format",
    "Cfs": "Chromatic Focal Shift",
    "Cgl": "Convert Global To Local",
    "Chk": "System Check",
    "Clg": "Convert Local to Global",
    "Cls": "Coating List",
    "Cna": "Coating, Ret. vs. Angle",
    "Cng": "Convert to NSC Group",
    "Cnw": "Coating, Ret. vs. Wavelength",
    "Coa": "Convert Asphere Type",
    "Con": "Conjugate Surface Analysis",
    "Cpa": "Coating, Phase vs. Angle",
    "Cpw": "Coating, Phase vs. Wavelength",
    "Cra": "Coating, Refl vs. Angle",
    "Crw": "Coating, Refl. vs. Wavelength",
    "Csd": "NSC concatenate spectral source files",
    "Csf": "NSC convert to spectral source file",
    "Cta": "Coating, Tran. vs. Angle",
    "Ctw": "Coating, Tran. vs. Wavelength",
    "Dbl": "Make Double Pass",
    "Dim": "Partially Coherent Image Analysis",
    "Dip": "Biocular Dipvergence/Convergence",
    "Dis": "Dispersion Plot",
    "Drs": "NSC download radiant source",
    "Dvl": "Dispersion vs. Wavelength Plot",
    "Dvr": "Detector Viewer",
    "Eca": "Explode CAD assembly",
    "Ect": "Edit Coating",
    "Ecp": "Explode Creo parametric assembly",
    "EDE": "Extra Data Editor",
    "Eec": "Export Encrypted Coating",
    "Eia": "Explode inventor assembly",
    "Ele": "ZEMAX Element Drawing",
    "Enc": "Diff Encircled Energy",
    "Esa": "Explode solidworks assembly",
    "Ext": "Exit",
    "Fba": "Find Best Asphere",
    "Fcd": "Field Curv/Distorion",
    "Fcl": "Fiber Coupling",
    "Fie": "Field Data",
    "Fld": "Add Fold Mirror",
    "Flx": "Delete Fold Mirror",
    "Fmm": "FFT MTF Map",
    "Foa": "Foucault Analysis",
    "Foo": "Footprint Analysis",
    "Fov": "Biocular Field of View Analysis",
    "Fps": "FFT PSF",
    "Fvw": "Flux vs. Wavelength",
    "Gbp": "Parax Gaussian Beam",
    "Gbs": "Skew Gaussian Beam",
    "Gcp": "Glass Compare",
    "Gee": "Geometric Encircled Energy",
    "Gen": "General Lens Data",
    "Gft": "Glass Fit",
    "Gho": "Ghost Focus",
    "Gip": "Grin Profile",
    "Gla": "Glass Catalog",
    "Glb": "Global Optimization",
    "Gmf": "Generate MAT file",
    "Gmm": "Geometric MTF Map",
    "Gmp": "Glass Map",
    "Gpr": "Gradium Profile",
    "Grd": "Grid Distortion",
    "Grs": "NSC generate radiant source",
    "Gst": "Glass Substitution Template",
    "Gtf": "Geometric MTF",
    "Gvf": "Geometric MTF vs. Field",
    "Ham": "Hammer Optimization",
    "Hcs": "Huygens PSF Cross Section",
    "Hlp": "Help",
    "Hmf": "Huygens MTF",
    "Hmh": "Huygens MTF vs Field",
    "Hps": "Huygens PSF",
    "Hsm": "Huygens Surface MTF",
    "Htf": "Huygens Through Focus MTF",
    "ISO": "ISO Element Drawing",
    "Ibm": "Geometric Bitmap Image Analysis",
    "Igs": "Export IGES File",
    "Iht": "Incident Angle vs. Image Height",
    "Ilf": "Illumination Surface",
    "Ils": "Illumination Scan",
    "Ima": "Geometric Image Analysis",
    "Imv": "IMA/BIM File Viewer",
    "Ins": "Insert Lens",
    "Int": "Interferogram",
    "Jbv": "Bitmap File Viewer",
    "L3d": "3D Layout",
    "L3n": "NSC 3D Layout",
    "LDE": "Lens Data Editor",
    "LSn": "NSC Shaded Model Layout",
    "Lac": "Last Configuration",
    "Lat": "Lateral Color",
    "Lay": "2D Layout",
    "Len": "Lens Search",
    "Lin": "Line/Edge Response",
    "Lok": "Lock All Windows",
    "Lon": "Longitudinal Aberration",
    "Lsa": "Light Source Analysis",
    "Lsf": "FFT Line/Edge Spread",
    "Lsh": "Shaded Model Layout",
    "Lsm": "Solid Model Layout",
    "Ltr": "NSC lighting trace",
    "Lwf": "Wireframe Layout",
    "MCE": "Multi-Config Editor",
    "MFE": "Merit Function Editor",
    "Mfl": "Merit Function List",
    "Mfo": "Make Focal",
    "Mtf": "Modulation Transfer Function (FFT MTF)",
    "Mth": "MTF vs. Field",
    "NCE": "Non-Sequential Editor",
    "New": "New File",
    "Nxc": "Next Configuration",
    "Obv": "NSC Object Viewer",
    "Opd": "Opd Fan",
    "Ope": "Open File",
    "Opt": "Optimization",
    "Pab": "Pupil Aberration Fan",
    "Pal": "Power Field Map",
    "Pat": "ZRD Path Analysis",
    "Pci": "Partially Coherent Image Analysis",
    "Pcs": "PSF Cross Section",
    "Per": "Performance Test",
    "Pha": "Pol. Phase Aberration",
    "Pmp": "Pol. Pupil Map",
    "Pol": "Pol. Ray Trace",
    "Pop": "Physical Optics Propagation",
    "Ppm": "Power Pupil Map",
    "Pre": "Prescription Data",
    "Prf": "Preferences",
    "Ptf": "Pol. Transmission Fan",
    "Pvr": "CAD part viewer",
    "Pzd": "Playback ZRD on Detectors",
    "Qad": "Quick Adjust",
    "Qfo": "Quick Focus",
    "Raa": "Remove All Apertures",
    "Ray": "Ray Fan",
    "Rcf": "Reload Coating File",
    "Rda": "NSC reverse radiance analysis",
    "Rdb": "Ray Database",
    "Rdw": "NSC roadway lighting analysis",
    "Red": "Redo",
    "Rel": "Relative Illumination",
    "Res": "Restore From Archive File",
    "Rev": "Reverse Elements",
    "Rfm": "RMS Field Map",
    "Rg4": "New Report Graphic 4",
    "Rg6": "New Report Graphic 6",
    "Rmf": "RMS vs. Focus",
    "Rml": "Refresh Macro List",
    "Rms": "RMS vs. Field",
    "Rmw": "RMS vs. Wavelength",
    "Rtc": "Ray Trace Control",
    "Rtr": "Ray Trace",
    "Rva": "Remove Variables",
    "Rxl": "Refresh Extensions List",
    "Sag": "Sag Table",
    "Sas": "Save As",
    "Sav": "Save File",
    "Sca": "Scale Lens",
    "Scc": "Surface Curvature Cross Section",
    "Scv": "Surface Curvature",
    "Sdi": "Seidel Diagram",
    "Sdv": "Src directivity",
    "Sei": "Seidel Coefficients",
    "Sff": "Full Field Spot",
    "Sfv": "Scatter Function Viewer",
    "Sim": "Image Simulation",
    "Sld": "Slider",
    "Slm": "Stock lens matching",
    "Sma": "Spot Matrix",
    "Smc": "Spot Matrix Config",
    "Smf": "Surface MTF",
    "Spc": "Surface Phase Cross Section",
    "Spj": "Src projection",
    "Spo": "Src polar",
    "Spt": "Spot Diagram",
    "Spv": "Scatter Polar Plot",
    "Srp": "Surface Phase",
    "Srs": "Surface Sag",
    "Ssc": "Surface Sag Cross Section",
    "Ssg": "System Summary Graphic",
    "Srv": "Src rms viewer",
    "Ssp": "src spectrum",
    "Stf": "Though Focus Spot",
    "Sti": "NSC convert SDF to IES",
    "Sur": "Surface Data",
    "Sys": "System Data",
    "TDE": "Tolerance Data Editor",
    "Tde": "Tilt/Decenter Elements",
    "Tfg": "Through Focus GTF",
    "Tfm": "Through Focus MTF",
    "Tls": "Tolerance List",
    "Tol": "Tolerancing",
    "Tpf": "Test Plate Fit",
    "Tpl": "Test Plate Lists",
    "Tra": "Pol. Transmission",
    "Trw": "Transmission vs. Wavelength",
    "Tsm": "Tolerance Summary",
    "Un2": "Universal Plot 2D",
    "Und": "Undo",
    "Uni": "Universal Plot",
    "Unl": "Unlock All Windows",
    "Upa": "Update All",
    "Upd": "Update",
    "Vig": "Vignetting Plot",
    "Vop": "Visual Optimization",
    "Vra": "Make All Radii Variable",
    "Vth": "Make All Thickness Variable",
    "Wav": "Wavelength Data",
    "Wfm": "Wavefront Map",
    "Xdi": "Extended Diffraction Image Ana",
    "Xis": "Export IGES/STEP/SAT FIle",
    "Xse": "Extended Source Encircled",
    "Yni": "YNI Contributions",
    "Yyb": "Y-Ybar",
    "Zat": "Zernike Annular Terms",
    "Zbb": "Export Zemax Black Box Data",
    "Zex": "ZEMAX Extensions",
    "Zfr": "Zernike Fringe Terms",
    "Zpd": "Zemax Part Designer",
    "Zpl": "Edit/Run ZPL Macros",
    "Zst": "Zernike Standard Terms",
    "Zvf": "Zernike Coefficients vs. Field"
    }

def showZButtonList():
    """List all the button codes

    showZButtonList()->None (the button codes are printed on screen)

    """
    print("Listing all ZEMAX Button codes:")
    for code, description in sorted(_Buttons.button_code.items()):
        _print_mod(_prettifyCodeDesc(code,description))
    _print_mod(_boldifyText("Total number of buttons = ",str(len(_Buttons.button_code))))

def getZButtonCount():
    """Returns the total number of buttons

    getZButtonCount()->count

    """
    return len(_Buttons.button_code)

def isZButtonCode(buttonCode):
    """Returns True or False depending on whether the button code is a valid
    button code.

    isZButtonCode(buttonCode)->bool

    Parameters
    ----------
    buttonCode : (string) the 3-letter case-sensitive button code to validate

    Returns
    -------
    bool        : True if valid button code, False otherwise
    """
    return str(buttonCode) in _Buttons.button_code.keys()

def showZButtonDescription(buttonCode):
    """Get a short description about a button code.

    showZButtonDescription(buttonCode)->description

    Parameters
    ----------
    buttonCode : (string) a 3-letter button code

    Returns
    -------
    description : a shot description about the button code function/analysis type.
    """
    if isZButtonCode(str(buttonCode)):
        _print_mod(_prettifyText(str(buttonCode), " is a ZEMAX button code",
                               color0='magenta',color1='black'))
        _print_mod(_prettifyText("Description: ", _Buttons.button_code[str(buttonCode)],
                  color0='blue',color1='black'))
    else:
        print("{} is NOT a valid ZEMAX button code.".format(str(buttonCode)))

def findZButtonCode(keywords):
    """Find Zemax button codes using specific keywords of interest.

    findZButtonCode("keyword#1 [, keyword#2, keyword#3, ...]")->searchResult

    Parameters
    ----------
    keywords : a string containing a list of comma separated keywords.

    Example
    -------
    >>> zb.findZButtonCode("Zernike")
    [Zst] Zernike Standard Terms
    [Zvf] Zernike Coefficients vs. Field
    [Zfr] Zernike Fringe Terms
    [Zat] Zernike Annular Terms

    Found 4 Button codes.

    >>> zb.findZButtonCode("Fan")
    [Opd] Opd Fan
    [Ray] Ray Fan
    [Ptf] Pol. Transmission Fan
    [Pab] Pupil Aberration Fan

    Found 4 Button codes.
    """
    words2find = [words.strip() for words in keywords.split(",")]
    previousFoundKeys = []
    for button, description in _Buttons.button_code.items():
        for kWord in words2find:
            if __find(kWord,description):
                _print_mod(_prettifyCodeDesc(button,description))
                previousFoundKeys.append(button)
                break # break the inner for loop
    if previousFoundKeys:
        _print_mod(_boldifyText("Found ", str(len(previousFoundKeys)),
                              " Button codes",'blue','red','blue'))

def __find(word2find,instring):
    r = _re.compile(r'\b({0})s?\b'.format(word2find),flags=_re.IGNORECASE)
    if r.search(instring):
        return True
    else:
        return False

if __name__ == '__main__':
    #Test showZButtonList()
    showZButtonList()
    #Test getZButtonCount()
    print("Total number of buttons:",getZButtonCount())
    #Test isZButtonCode()
    print("'Pre' is a button code (True/False):", isZButtonCode('Pre'))
    print("'Wav' is an button code (True/False):", isZButtonCode('Wav'))
    print("'RANDOM' is a button code (True/False):", isZButtonCode('RANDOM'))
    #Test showZOperandDescription()
    showZButtonDescription('RANDOM')
    showZButtonDescription('Vth')


