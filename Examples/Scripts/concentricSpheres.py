# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:      concentricSpheres.py
# Purpose:   demonstrates the use of arraytrace module 
#
# Author:     Julian Stuermer
#
# Created:    June 2015
# Copyright:  (c) Julian Stuermer, 2012 - 2015
# Licence:    MIT License
#-------------------------------------------------------------------------------
from __future__ import print_function, division
import os
import pyzdde.arraytrace as at # Module for array ray tracing
import pyzdde.zdde as pyz
import matplotlib.pyplot as plt
from mpl_toolkits.axes_grid1 import make_axes_locatable
import numpy as np
from scipy.interpolate import griddata


def handle_figure(fig, name, bsavefig):
    '''helper function to save/ display matplotlib figure'''
    if bsavefig:
        fig.savefig(name, dpi=100, bbox_inches='tight', pad_inches=0.1)
    else:
        plt.show()

def full_field_spot_diagramm(N=51, zoom=75, box=25, bsavefig=False):
    '''creates spot diagramm over the full field, similar to Zemax 
    '''
    print('Computing Full Field Spot Diagramm')
    # Define rays
    nx = ny = N # number of rays in x- and y-direction
    px = np.repeat(np.linspace(-1, 1, nx), nx)
    py = np.tile(np.linspace(-1, 1, ny), ny)
    r = np.sqrt(px**2 + py**2)
    numRays = nx*ny
    # Get field data and create Plot
    fieldData = ln.zGetField(0)   
    fig, ax = plt.subplots(figsize=(fieldData.maxX*2, fieldData.maxY*2))
    fields = ln.zGetFieldTuple()
    # Box around spots in mm
    box_size = box / 1000
    for field in fields:
        hx = np.repeat(field.xf / fieldData.maxX, numRays)
        hy = np.repeat(field.yf / fieldData.maxY, numRays)
        rd = at.zGetTraceArray(numRays, px=px.tolist(), py=py.tolist(), 
                               hx=hx.tolist(), hy=hy.tolist(), waveNum=1)
        # remove vignetted rays and rays outside a circular pupil (r>1)
        vig = np.array(rd[1])
        valid_rays = np.logical_and(vig < 1, r <= 1)
        x = np.array(rd[2])[valid_rays]
        y = np.array(rd[3])[valid_rays]
        xMean, yMean = np.mean(x), np.mean(y)
        x_c, y_c = x - xMean, y - yMean
        
        ax.scatter(xMean + zoom * x_c, yMean + zoom * y_c, s=0.3, c='blue', lw=0)    
        box = plt.Rectangle((xMean - box_size/2 * zoom, yMean - box_size/2 * zoom),
                             zoom * box_size, zoom * box_size, fill=False)
        ax.add_patch(box)
    ax.grid()
    ax.set_title('Full Field Spot Diagram')
    ax.set_xlabel('Detector x [mm]')
    ax.set_ylabel('Detector y [mm]')
    handle_figure(fig, 'fullfield.png', bsavefig)


def spots_matrix(nFields=5, N=51, zoom=75., box=25, bsavefig=False):
    '''create spot diagrams for nFields x nFields ray color by pupil position
    '''
    print('Computing Spots Matrix')
    # Define rays
    nx = ny = N # number of rays in x- and y-direction
    px = np.repeat(np.linspace(-1, 1, nx), nx)
    py = np.tile(np.linspace(-1, 1, ny), ny)
    r = np.sqrt(px**2 + py**2)
    numRays = nx*ny
    fields = [(x, y) for y in np.linspace(1, -1, nFields) 
                        for x in np.linspace(-1, 1, nFields)]
    ENPD = ln.zGetPupil().ENPD
    fig, ax = plt.subplots(figsize=(10, 10))
    
    # box around spots in mm
    box_size = box / 1000
    for i, field in enumerate(fields):
        hx = np.repeat(field[0], numRays)
        hy = np.repeat(field[1], numRays)
        rd = at.zGetTraceArray(numRays, px=px.tolist(), py=py.tolist(), 
                               hx=hx.tolist(), hy=hy.tolist(), waveNum=1)
        # remove vignetted rays and rays outside a circular pupil (r>1)
        vig = np.array(rd[1])
        valid_rays = np.logical_and(vig < 1, r <= 1)
        x = np.array(rd[2])[valid_rays]
        y = np.array(rd[3])[valid_rays]
        xMean = np.mean(x)
        yMean = np.mean(y)
        x_c = x - xMean
        y_c = y - yMean
        spotSTD = np.std(np.sqrt(x_c**2 + y_c**2))
        adaptivezoom = zoom/(spotSTD*100)
        box = plt.Rectangle((xMean - box_size/2 * adaptivezoom, 
                             yMean - box_size/2 * adaptivezoom),
                             adaptivezoom * box_size,
                             adaptivezoom * box_size, fill=False)
        ax.add_patch(box)
        scatter = ax.scatter(xMean + adaptivezoom * x_c, yMean + adaptivezoom * y_c, 
                             s=0.5, c=r[valid_rays]*ENPD, cmap=plt.cm.jet, lw=0)
    ax.set_ylim((-20, 20))
    ax.set_xlim((-20, 20))
    ax.set_xlabel('Detector x [mm]')
    ax.set_ylabel('Detector y [mm]')
    divider = make_axes_locatable(ax)
    cax = divider.append_axes("right", size="5%", pad=0.05)
    cb = fig.colorbar(scatter, cax=cax)
    cb.set_label('Pupil Diameter [mm]')
    ax.set_title('Adaptive Zoom Spots / Pupil size')
    ax.set_aspect('equal')
    fig.tight_layout()
    handle_figure(fig, 'coloredpupil.png', bsavefig)

    
def spot_radius_map(nFields= 15, N=51, EE=80, zoom=10, bsavefig=False):
    '''
    create spot diagrams for nFields x nFields 
    calculate spot radii and generate density map
    '''
    print('Computing Spot Radius Map')
    # Define rays
    nx = ny = N # number of rays in x- and y-direction
    px = np.repeat(np.linspace(-1, 1, nx), nx)
    py = np.tile(np.linspace(-1, 1, ny), ny)
    r = np.sqrt(px**2 + py**2)
    numRays = nx*ny
    fields = [(x,y) for y in np.linspace(1, -1, nFields) 
                       for x in np.linspace(-1, 1, nFields)]
    
    fig, ax = plt.subplots(figsize=(10, 10))
    X, Y, X_MEAN, Y_MEAN, R = [], [], [], [], []
    for i, field in enumerate(fields):
        hx = np.repeat(field[0], numRays)
        hy = np.repeat(field[1], numRays)    
        rd = at.zGetTraceArray(numRays, px=px.tolist(), py=py.tolist(), 
                               hx=hx.tolist(), hy=hy.tolist(), waveNum=1)
        # remove vignetted rays and rays outside a circular pupil (r>1)
        vig = np.array(rd[1])
        valid_rays = np.logical_and(vig < 1, r <= 1)
        x = np.array(rd[2])[valid_rays]
        y = np.array(rd[3])[valid_rays]
        xMean, yMean = np.mean(x), np.mean(y)
        x_c, y_c = x - xMean, y - yMean
        radius = np.sqrt(x_c**2 + y_c**2)
        X.append(xMean)
        Y.append(yMean)
        X_MEAN.append(xMean)
        Y_MEAN.append(yMean)
        R.append(np.sort(radius)[int(EE/100*len(radius))]*1000)
        ax.scatter(xMean + zoom * x_c, yMean + zoom * y_c, s=0.2, c='k', lw=0)        
    X, Y = np.array(X), np.array(Y)
    xi = np.linspace(np.min(X), np.max(X), 50)
    yi = np.linspace(np.min(Y), np.max(Y), 50)
    zi = griddata((X, Y), np.array(R), (xi[None,:], yi[:,None]), method='cubic')
    ax.set_xlim(np.min(xi), np.max(xi))
    ax.set_ylim(np.min(yi), np.max(yi))
    ax.set_xlabel('Detector x [mm]')
    ax.set_ylabel('Detector y [mm]')
    ax.set_title('Encircled Energy Map')
    im = ax.imshow(zi, interpolation='bilinear', cmap=plt.cm.jet,
                   extent=[np.min(xi), np.max(xi), np.min(yi), np.max(yi)])
    divider = make_axes_locatable(ax)
    cax = divider.append_axes("right", size="5%", pad=0.05)
    cb = plt.colorbar(im, cax=cax)
    cb.set_label('EE' + str(EE) + ' radius [micron]')
    fig.tight_layout()
    handle_figure(fig, 'map.png', bsavefig)

if __name__ == '__main__':
    cd = os.path.dirname(os.path.realpath(__file__))
    ind = cd.find('Examples')  
    pDir = cd[0:ind-1]
    zmxfile = 'ConcentricSpheres.zmx'
    filename = os.path.join(pDir, 'ZMXFILES', zmxfile) # Change the filename appropriately 
    ln = pyz.createLink() 
    ln.zLoadFile(filename)
    ln.zPushLens() # push lens to the LDE as array tracing only works with the file in the LDE
    full_field_spot_diagramm(bsavefig=False)
    spots_matrix(bsavefig=False)
    spot_radius_map(bsavefig=False)
    ln.close()
    print('Done')