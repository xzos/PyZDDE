from zdde import readInBeamFile
import os
import matplotlib.pyplot as plt
directory = os.path.dirname(os.path.realpath(__file__))

beamfilename = directory+"\\type2spdc.zbf"

(version,(nx,ny),ispol,units,(dx,dy),(zposition_x,zposition_y),(rayleigh_x,rayleigh_y),(waist_x,waist_y),lambd,index,(x_matrix,y_matrix),(Ex_real,Ex_imag,Ey_real,Ey_imag)) = readInBeamFile(beamfilename)

xlabels = [-nx*dx/2+x*dx for x in range(0,nx)]
ylabels = [-ny*dy/2+y*dy for y in range(0,ny)]

fig, ax = plt.subplots()

ax.contourf(xlabels,ylabels,Ey_real,cmap="Blues",alpha=0.66,label="Y polarization")
ax.contourf(xlabels,ylabels,Ex_real,cmap="Reds",alpha=0.66,label="X polarization")

ax.legend(loc='upper center', shadow=True)
ax.set_xlabel("X position (mm)")
ax.set_ylabel("Y position (mm)")
plt.show()