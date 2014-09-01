#-------------------------------------------------------------------------------
# Name:     beamFileWrite.py
# Purpose:  Demonstrate the function writeBeamFile()
#           The example uses the python image library (PIL) to transfer the
#           information from a PNG to a zemax beam file format. You might do
#           this, for example, if you have an image from a camera in your setup
#           of the rings created by parametric down conversion that you want to
#           use in your zemax ray tracing. I've chosen to use a recognizable
#           image for the example so that it can be confirmed that it is
#           working as expected.
#
# NOTE:      Please note that this code uses PIL (Python Image Library) and
#            matplotlib
#
# Copyright: (c) 2012 - 2014
# Licence:   MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function
from PIL import Image
from pyzdde.zdde import writeBeamFile, readBeamFile
import matplotlib.pyplot as plt
import os
import time

directory = os.path.dirname(os.path.realpath(__file__))
im = Image.open(directory+os.path.sep+"pikachu2.png")

pix = im.load()
#(nx, ny)  = im.size
(nx, ny) = (64, 64)

#
Ex_real = [[0 for x in range(nx)] for y in range(ny)]
Ex_imag = [[0 for x in range(nx)] for y in range(ny)]
Ey_real = [[0 for x in range(nx)] for y in range(ny)]
Ey_imag = [[0 for x in range(nx)] for y in range(ny)]

for i in range(ny):
    for j in range(nx):
        Ex_real[nx-j-1][i] = pix[i, j]

# ATTENTION: imshow will show a flipped image. Nevertheless the image will be in correct orientation in zemax.
plt.imshow(Ex_real)
plt.show()

n=(nx,ny)
efield = (Ex_real, Ex_imag, Ey_real, Ey_imag)
version = 0
units = 0
ispol = True
d = (0.625,0.625)
zposition = (0.0,0.0)
rayleigh = (0.0,0.0)
waist = (3.0,3.0)
lamda = 0.00055
index = 1.0
receiver_eff = 0
system_eff = 0

beamfilename = directory+os.path.sep+"pikachu2.zbf"

# Write the beam file
writeBeamFile(beamfilename, version, n, ispol, units, d, zposition,
              rayleigh, waist, lamda, index, efield)

print("Done writing the beam file")
time.sleep(0.5)

# Read the beamfile just created and display
beamData = readBeamFile(beamfilename)

(version, (nx, ny), ispol, units, (dx,dy), (zposition_x, zposition_y),
 (rayleigh_x, rayleigh_y), (waist_x, waist_y), lamda, index, receiver_eff, system_eff,
 (x_matrix, y_matrix), (Ex_real, Ex_imag, Ey_real, Ey_imag)) = beamData

xlabels = [-nx*dx/2+x*dx for x in range(0, nx)]
ylabels = [-ny*dy/2+y*dy for y in range(0, ny)]

fig, ax = plt.subplots()
cs = ax.contourf(xlabels, ylabels, Ex_real, cmap="Blues", alpha=0.66)
ax.set_aspect('equal')
ax.set_xlabel("X position (mm)")
ax.set_ylabel("Y position (mm)")

print("\n")
print("************************")
print("Data read from beam file")
print("************************")
print("Version: ", version)
print("n :", (nx, ny))
print("ispol :", ispol)
print("units :", units)
print("d :", (dx, dy))
print("zposition :", (zposition_x, zposition_y))
print("rayleigh :", (rayleigh_x, rayleigh_y))
print("waist :", (waist_x, waist_y))
print("lamba :", lamda)
print("index :", index)
plt.show()
