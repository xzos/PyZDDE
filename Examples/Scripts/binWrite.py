import Image
from zdde import writeBeamFile
import os
directory = os.path.dirname(os.path.realpath(__file__))
im = Image.open(directory+"\pikachu2.png")
pix =im.load()
#(nx,ny)  = im.size
(nx,ny) = (64,64)

Ex_real = [[0 for x in xrange(ny)] for x in xrange(nx)]
Ex_imag = [[0 for x in xrange(ny)] for x in xrange(nx)]
Ey_real = [[0 for x in xrange(ny)] for x in xrange(nx)]
Ey_imag = [[0 for x in xrange(ny)] for x in xrange(nx)]        

for i in range(nx):
    for j in range(ny):
        Ex_real[ny-j-1][i] = pix[i,j]
        #print(pix[i,j])

n=(nx,ny)
efield = (Ex_real,Ex_imag,Ey_real,Ey_imag)
version = 0
units = 0
ispol = True
d = (0.625,0.625)
zposition = (0.0,0.0)
rayleigh = (0.0,0.0)
waist = (3.0,3.0)
lambd = 0.00055
index = 1.0

beamfilename = directory+"\pikachu2.zbf"

writeBeamFile(beamfilename, version, n, ispol,units,d,zposition,rayleigh,waist,lambd,index,efield)
        