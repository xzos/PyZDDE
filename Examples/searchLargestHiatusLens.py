#-------------------------------------------------------------------------------
# Name:        searchLargestHiatusLens.py
# Purpose:     Search for a specific attribute through ZMX files. In this program
#              we are interested in finding the lens (design) that has the largest
#              (or smallest) hiatus also called nodal space, Null space, or the
#              interstitium (i.e. the distance between the two principal planes)
#
#              Note [from Zemax Manual]:
#              Object space positions are measured with respect to surface 1.
#              Image space positions are measured with respect to the image surface.
#              The index in both the object space and image space is considered.
#
#              Assumptions:
#              1. The last surface is the image surface
#              2. Search only sequential zemax files. The prescription file output
#                 for Non-sequential Zemax analysis is different from sequential.
#              3. File-names within the search directory are unique.
#
#              Note:
#              1. If Zemax is unable to open a .zmx certain file, it pops-up an
#                 error msg, which the user needs to click. So, in such scenarios
#                 this program execution would be stalled until the user has clicked
#                 on the message. That particular file is then excluded from the
#                 analysis.
#
# Author:      Indranil Sinharoy, Southern Methodist University
#
# Created:     23/04/2013
# Copyright:   (c) Indranil Sinharoy, Southern Methodist University, 2013
# Licence:     MIT License
#-------------------------------------------------------------------------------

from __future__ import division
from __future__ import print_function
import os, glob, sys, fnmatch
from operator import itemgetter
import Tkinter, tkFileDialog, Tkconstants
import datetime

# Put both the "Examples" and the "PyZDDE" directory in the python search path.
exampleDirectory = os.path.dirname(os.path.realpath(__file__))
ind = exampleDirectory.find('Examples')
pyzddedirectory = exampleDirectory[0:ind-1]
if exampleDirectory not in sys.path:
    sys.path.append(exampleDirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

import pyzdde

#Program control parameters
ORDERED_HIATUS_DATA_IN_FILE = True   # Sorted output in a file ? [May take a little longer time]
SCALE_LENSES = True                  # Scale lenses/Normalize all lenses to
NORMALIZATION_EFL = 500.00           # Focal length to use for Normalization
ORDERING   = 'large2small'           # 'large2small' or 'small2large'
HIATUS_UPPER_LIMIT = 20000.00        # Ignore lenses for which hiatus is greater than some value
fDBG_PRINT = False                   # Turn off/on the debug prints

# ZEMAX file DIRECTORY to search (can have sub-directories)
zmxfp = pyzddedirectory+"\\ZMXFILES"
#A simple Tkinter GUI prompting for directory
root = Tkinter.Tk()
class TkFileDialog(Tkinter.Frame):
    def __init__(self, root):
        Tkinter.Frame.__init__(self, root, borderwidth=20,height=32,width=42)

        #Top-level label
        self.label0 = Tkinter.Label(self,text = "Find eXtreme Hiatus",
                           font=("Helvetica",16),fg='blue',justify=Tkinter.LEFT)
        self.label0.pack()

        # options for buttons
        button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}
        checkBox_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

        # define first button
        self.b1 = Tkinter.Button(self, text='Select Directory', command=self.askdirectory)
        self.b1.pack(**button_opt)

        #Add a checkbox button (for lens scaling option)
        self.lensScaleOptVar = Tkinter.IntVar(value=0)
        self.c1 = Tkinter.Checkbutton(self,text="Enable Lens scaling ?",
                 variable=self.lensScaleOptVar,command=self.cb1,onvalue=1)
        self.c1.pack(**checkBox_opt)
        self.c1.select()  #The check-box is checked initially

        #Add a label to indicate/enter normalization EFL
        self.label1 = Tkinter.Label(self,text = "Normalization EFL", justify=Tkinter.LEFT)
        self.label1.pack()

        #Add Entry Widget to enter default normalization EFL
        self.normEFLVar = Tkinter.StringVar()
        self.normEFLentry = Tkinter.Entry(self,text="test",textvariable=self.normEFLVar)
        self.normEFLentry.pack()
        self.normEFLentry.insert(0,str(NORMALIZATION_EFL))

        #Add another label
        self.label2 = Tkinter.Label(self,text = "Ignore values above:", justify=Tkinter.LEFT)
        self.label2.pack()

        #Add an Entry Widget to enter value for upper level hiatus (string)
        self.maxHiatusVar = Tkinter.StringVar()
        self.maxHiatusEntry = Tkinter.Entry(self,text="test",textvariable=self.maxHiatusVar)
        self.maxHiatusEntry.pack()
        self.maxHiatusEntry.insert(0,str(HIATUS_UPPER_LIMIT))

        # checkbox button 2 (For text dump option)
        self.txtFileDumpVar = Tkinter.IntVar(value=0)
        self.c2 = Tkinter.Checkbutton(self,text="Save to a TXT file?",
                                 variable=self.txtFileDumpVar,command=self.cb2,onvalue=1)
        self.c2.pack(**checkBox_opt)
        self.c2.select()   #The check-box is checked initially

        #Add a "Find" button
        self.b2 = Tkinter.Button(self,text='Find',fg="red",command=self.find)
        self.b2.pack(**button_opt)

    def askdirectory(self):
        """Returns a selected directoryname."""
        global zmxfp
        zmxfp = tkFileDialog.askdirectory(parent=root,initialdir=zmxfp,
                            title='Please navigate to a directory')
        return

    def cb1(self):
        global SCALE_LENSES
        SCALE_LENSES = bool(self.lensScaleOptVar.get())
        if ~SCALE_LENSES:
            #self.normEFLentry.
            pass
        return

    def cb2(self):
        global ORDERED_HIATUS_DATA_IN_FILE
        ORDERED_HIATUS_DATA_IN_FILE = bool(self.txtFileDumpVar.get())
        return

    def find(self):
        global HIATUS_UPPER_LIMIT
        global NORMALIZATION_EFL
        self.normEFLentry.focus_set()
        NORMALIZATION_EFL = float(self.normEFLentry.get())
        self.maxHiatusEntry.focus_set()
        HIATUS_UPPER_LIMIT = float(self.maxHiatusEntry.get())
        root.quit()
        root.destroy()

TkFileDialog(root).pack()
root.mainloop()
#end of Tikinter GUI code

# Create a DDE channel object
pyZmLnk = pyzdde.PyZDDE()
#Initialize the DDE link
stat = pyZmLnk.zDDEInit()

#Get all the zemax files in the directories recursively
pattern = "*.zmx"
filenames = [os.path.join(dirpath,f)
             for dirpath, subFolders, files in os.walk(zmxfp)
             for f in fnmatch.filter(files,pattern)]
parentFolder = str(os.path.split(zmxfp)[1])

###To just use one file FOR DEBUGGING PURPOSE -- comment out this section
##oneFile = []
##oneFile.append(filenames[1])
##filenames = oneFile
###end of "just use one file to test"

print("SCALE_LENSES: ", SCALE_LENSES)
print("NORMALIZATION_EFL: ", NORMALIZATION_EFL)


now = datetime.datetime.now()

# ###################
# MAIN CODE LOGIC
# ###################
#Create a dictionary to store the filenames and hiatus
hiatusData = dict()
scaleFactorData = dict()
largestHiatusValue =   0.0     #init the variables for largest hiatus
largestHiatusLensFile = "None"
lensFileCount = 0
totalNumLensFiles = len(filenames)
totalFilesNotLoaded = 0 #File count of files that couldn't be loaded by Zemax
filesNotLoaded = []   #List of files that couldn't be loaded by Zemax
# Loop through all the files in filenames, load the zemax files, get the data
for lens_file in filenames:
    if fDBG_PRINT:
        print("Lens file: ",lens_file)
    #Load the lens in to the Zemax DDE server
    ret = pyZmLnk.zLoadFile(lens_file)
    if ret != 0:
        print(ret, lens_file, " Couldn't open!")
        filesNotLoaded.append(lens_file)
        totalFilesNotLoaded +=1
        continue
    #assert ret == 0
    #In order to maintain the units, set the units to mm for all lenses. Also
    #ensure that the global reference surface for all lenses is set to surface 1,
    #all other system settings should remain same.
    recSystemData_g = pyZmLnk.zGetSystem() #Get the current system parameters
    numSurf       = recSystemData_g[0]
    unitCode      = recSystemData_g[1]  # lens units code (0,1,2,or 3 for mm, cm, in, or M)
    stopSurf      = recSystemData_g[2]
    nonAxialFlag  = recSystemData_g[3]
    rayAimingType = recSystemData_g[4]
    adjust_index  = recSystemData_g[5]
    temp          = recSystemData_g[6]
    pressure      = recSystemData_g[7]
    globalRefSurf = recSystemData_g[8]
    #Set the system parameters
    recSystemData_s = pyZmLnk.zSetSystem(0,stopSurf,rayAimingType,0,temp,pressure,1)

    #Scale lens to a normalized EFFL
    scaleFactor = 1.00
    if SCALE_LENSES:
        #Get first order EFL
        efl = pyZmLnk.zGetFirst()[0]
        #Determine scale factor
        scaleFactor = abs(NORMALIZATION_EFL/efl)
        if fDBG_PRINT:
            print("EFFL: ",efl," Scale Factor: ", scaleFactor)
        #Scale Lens
        ret_ls = pyZmLnk.zLensScale(scaleFactor)

        if ret_ls == -1:  # Lens scale failure, don't bother to calculate hiatus
            print("Lens scaling failed for: ",lens_file)
            continue

    #Update the lens
    #ret = pyZmLnk.zGetUpdate() ... I don't think the designs should be updated...
    #as we don't need to re-optimize, etc.
    #assert ret == 0
    textFileName = exampleDirectory + '\\' + "searchSpecAttr_Prescription.txt"

    #Get the Hiatus for the lens design
    hiatus = pyZmLnk.zCalculateHiatus(textFileName,keepFile=False)

    if hiatus > HIATUS_UPPER_LIMIT:
        continue
    lensFileCount +=1  #Increment the lens files count
    if hiatus > largestHiatusValue:
        largestHiatusValue = hiatus
        largestHiatusLensFile = os.path.basename(lens_file)


    #Add to the dictionary
    hiatusData[os.path.basename(lens_file)] = hiatus
    scaleFactorData[os.path.basename(lens_file)] = scaleFactor

#Close the DDE channel before processing the dictionary.
pyZmLnk.zDDEClose()

if fDBG_PRINT:
    print("Hiatus data dictionary:\n", hiatusData)

if ORDERED_HIATUS_DATA_IN_FILE:
    #Sort the "dictionary" in 'large2small' or 'small2large' order
    #The output (hiatusData_sorted) is a list of tuples
    if ORDERING == 'small2large':
        hiatusData_sorted = sorted(hiatusData.items(),key=itemgetter(1))
    else:
        hiatusData_sorted = sorted(hiatusData.items(),key=itemgetter(1),reverse=True)
    #Open a file for writing the data
    dtStamp = "_%d_%d_%d_%dh_%dm_%ds" %(now.year,now.month,now.day,now.hour,now.minute,now.second)
    fileref_hds = open("searchLargestHiatusLens_"+parentFolder+dtStamp+".txt",'w')
    fileref_hds.write("LENS HIATUS MEASUREMENT:\n\n")
    fileref_hds.write("Date and time: " + now.strftime("%Y-%m-%d %H:%M"))
    fileref_hds.write("\nUnits: mm")
    if SCALE_LENSES:
        fileref_hds.write("\nLens Scaling for normalization: ON. Normalization EFL = %1.2f"%(NORMALIZATION_EFL))
    else:
        fileref_hds.write("\nLens Scaling for normalization: OFF")
    fileref_hds.write("\nDirectory: "+ zmxfp)
    fileref_hds.write("\n%s Lenses analyzed out of %s lenses!"%(lensFileCount,
                                                        totalNumLensFiles))
    fileref_hds.write("\nLens files not loaded by Zemax: %s (See list below)"%(totalFilesNotLoaded))
    fileref_hds.write("\nLenses with hiatus above %s have been ignored.\n\n"%(HIATUS_UPPER_LIMIT))
    fileref_hds.write("\nThe sorted list is:\n\n")
    for i in hiatusData_sorted:
        fileref_hds.write("%s\t\t%1.2f\t(scale factor = %1.2f)\n"%(i[0],i[1],scaleFactorData[i[0]]))
    fileref_hds.write("\n\nLens files that Zemax couldn't open for analysis:\n\n")
    for fl in filesNotLoaded:
        fileref_hds.write("%s\n"%fl)
    fileref_hds.close()

#Print the largest lens having the largest hiatus and the hiatus value
print(lensFileCount, "lenses analyzed for largest hiatus (in mm) out of", totalNumLensFiles, "lenses.")
print("Largest Hiatus Lens:", largestHiatusLensFile)
print("Hiatus:", largestHiatusValue)
