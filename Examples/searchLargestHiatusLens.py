#-------------------------------------------------------------------------------
# Name:        searchLargestHiatusLens.py
# Purpose:     Search for a specific attribute through ZMX files. In this program
#              we are interested in finding the lens (design) that has the largest
#              (or smallest) hiatus or nodal space or interstitium (i.e. the
#              distance between the two principal planes)
#
#              Note [from Zemax Manual]:
#              Object space positions are measured with respect to surface 1.
#              Image space positions are measured with respect to the image surface.
#              The index in both the object space and image space is considered.
#
#              Assumptions:
#              1. The last surface is the image surface
#              2.
#
#              Note:
#              1. If Zemax is unable to open a .zmx certain file, it pops-up an
#                 error msg, which the user needs to click. So, in such scenarios
#                 this program execution would be stalled until the user has clicked
#                 on the message. That particular file is then excluded from the
#                 analysis.
#
# Author:      Indranil Sinharoy
#
# Created:     23/04/2013
# Copyright:   (c) Indranil 2013
# Licence:     MIT License
#-------------------------------------------------------------------------------

from __future__ import division
from __future__ import print_function
import os, glob, sys, fnmatch
from operator import itemgetter
import Tkinter, tkFileDialog
import datetime

# Put both the "Examples" and the "PyZDDE" directory in the python search path.
exampleDirectory = os.getcwd()
ind = exampleDirectory.find('Examples')
pyzddedirectory = exampleDirectory[0:ind-1]
if exampleDirectory not in sys.path:
    sys.path.append(exampleDirectory)
if pyzddedirectory not in sys.path:
    sys.path.append(pyzddedirectory)

import pyZDDE

#Program control parameters
ORDERED_HIATUS_DATA_IN_FILE = True   # Sorted output in a file ? [will take longer time]
ORDERING   = 'large2small'           # 'large2small' or 'small2large'
fDBG_PRINT = False                   # Turn off/on the debug prints

# ZEMAX file DIRECTORY to search (can have sub-directories)
zmxfp = pyzddedirectory+'\\ZMXFILES\\'

#A simple Tkinter GUI prompting for directory
root = Tkinter.Tk()
root.withdraw()
zmxfp = tkFileDialog.askdirectory(parent=root,initialdir=dirname,
                            title='Please navigate to a directory')
#end of Tikinter GUI code

# Create a DDE channel object
pyZmLnk = pyZDDE.pyzdde()
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

now = datetime.datetime.now()
#Create a dictionary to store the filenames and hiatus
hiatusData = dict()
largestHiatusValue =   0.0     #init the variables for largest hiatus
largestHiatusLensFile = "None"
lensFileCount = 0
# Loop through all the files in filenames, load the zemax files, get the data
for lens_file in filenames:
    if fDBG_PRINT:
        print("Lens file: ",lens_file)
    #Load the lens in to the Zemax DDE server
    ret = pyZmLnk.zLoadFile(lens_file)
    if ret != 0:
        print(ret, lens_file, " Couldn't open!")
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

    #Push the lens in the Zemax DDE server into the LDE
    #ret = pyZmLnk.zPushLens(updateFlag=1) #Not entirely sure if pushLens is required
    #Update the lens
    ret = pyZmLnk.zGetUpdate()
    assert ret == 0
    #Qucik focus/optimize?
    # I don't think it is required, as, the principal planes doesn't change with focusing.
    #Should I do Refresh() [ie copy the lens data from the LDE into the stored
    #copy of the Zemax server]
    #ret = pyZmLnk.zGetRefresh()
    #assert ret == 0
    #Dump the Prescription file
    textFileName = exampleDirectory + '\\' + "searchSpecAttr_Prescription.txt"
    ret = pyZmLnk.zGetTextFile(textFileName,'Pre',"None",0)
    assert ret == 0
    #Open the text file in read mode to read
    fileref = open(textFileName,"r")
    principalPlane_objSpace = 0.0; principalPlane_imgSpace = 0.0; hiatus = 0.0
    count = 0
    #The number of expected Principal planes in each Pre file is equal to the
    #number of wavelengths in the general settings of the lens design
    line_list = fileref.readlines()
    fileref.close()
    #See Endnote 1 for the reasons why the file was not read as an iterable object
    #and instead, we create a list of all the lines in the file, which is obviously
    #very wasteful of memory

    for line_num,line in enumerate(line_list):
        #Extract the image surface distance from the global ref sur (surface 1)
        sectionString = "GLOBAL VERTEX COORDINATES, ORIENTATIONS, AND ROTATION/OFFSET MATRICES:"
        if line.rstrip()== sectionString:
            ima_3 = line_list[line_num + numSurf*4 + 6]
            ima_z = float(ima_3.split()[3])
            if fDBG_PRINT:
                print("Image surface:", ima_z)
        #Extract the Principal plane distances.
        if "Principal Planes" in line and "Anti" not in line:
            principalPlane_objSpace += float(line.split()[3])
            principalPlane_imgSpace += float(line.split()[4])
            count +=1  #Increment (wavelength) counter for averaging
    lensFileCount +=1  #Increment the lens files count
   #Calculate the average (for all wavelengths) of the principal plane distances
    if count > 0:
        principalPlane_objSpace = principalPlane_objSpace/count
        principalPlane_imgSpace = principalPlane_imgSpace/count
        #Calculate the hiatus (only if count > 0) as
        #hiatus = (img_surf_dist + img_surf_2_imgSpacePP_dist) - objSpacePP_dist
        hiatus = abs(ima_z + principalPlane_imgSpace - principalPlane_objSpace)
    if fDBG_PRINT:
        print("Object space Principal Plane: ", principalPlane_objSpace)
        print("Image space Principal Plane: ", principalPlane_imgSpace)
        print("Hiatus: ", hiatus)

    if hiatus > largestHiatusValue:
        largestHiatusValue = hiatus
        largestHiatusLensFile = os.path.basename(lens_file)


    #Add to the dictionary
    hiatusData[os.path.basename(lens_file)] = hiatus

#Close the DDE channel before processing the dictionary.
pyZmLnk.zDDEClose()

#Delete the prescription file (the directory remains clean)
os.remove(textFileName)

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
    fileref_hds.write("Lens hiatus measurement:\n\n")
    fileref_hds.write("Date and time: " + now.strftime("%Y-%m-%d %H:%M"))
    fileref_hds.write("\nUnits: mm")
    fileref_hds.write("\nDirectory: "+ zmxfp)
    fileref_hds.write("\n%s Lenses Analyzed!\n\nThe sorted list is:\n\n"%(lensFileCount))
    for i in hiatusData_sorted:
        fileref_hds.write("%s\t\t%s\n"%(i[0],i[1]))
    fileref_hds.close()

#Print the largest lens having the largest hiatus and the hiatus value
print(lensFileCount, " lenses analyzed for largest hiatus (in mm): ")
print("Lens:", largestHiatusLensFile)
print("Hiatus:", largestHiatusValue)

#Endnote 1:
#It is very difficult (if not impossible) to read the prescirption files using bytes
#as we want to get to a specific position based on "keywords" and not "bytes". (we
#are not guaranteed to find the same "keyword" for a specific byte-based-position
#everytime we read a prescription file)
#If we the file line by line as "for line in file" (using the file iterable object)
#it is hard to read, identify and store a specific line which doesn't have any
#keywords. Also, because of possible dataloss, Python raises an exception, if we
#try to do readline() or readlines() within the "for line in file" iteration.