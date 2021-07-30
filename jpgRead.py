# -*- coding: utf-8 -*-
"""
Created on Thu Jul 29 13:02:54 2021

@author: brendanlia
"""

import cv2 as cv
import numpy as np
import glob
import matplotlib.pyplot as plt
import pandas as pd
import plotdigitizer
import time

# cv uses BGR while matplotlib uses RGB
# Load image as RGB 3D array

def jpgtoCsv(jpgName,csvName):
    img = cv.imread(jpgName)
    img = cv.cvtColor(img, cv.COLOR_BGR2RGB)
    # plt.imshow(img)
    
    # Find indexes of RGB values close to the line colours
    sensitivity = 5
    green = [158,171,117]
    blue = [104,125,162]
    purple = [126,113,140]
    red = [150,91,94]
    
    greenindex = []
    blueindex = []
    purpleindex = []
    redindex = []
    
    for i in range(len(img[1])):
        for j in range(len(img)):
            if img[j,i,0] == 255:
                pass
            elif img[j,i,0] < 70:
                pass
            elif np.average(abs(np.subtract(green,img[j,i,:]))) < sensitivity:
                coord = [i,j]
                greenindex.append(coord)
            elif np.average(abs(np.subtract(blue,img[j,i,:]))) < sensitivity:
                coord = [i,j]
                blueindex.append(coord)
            elif np.average(abs(np.subtract(purple,img[j,i,:]))) < sensitivity:
                coord = [i,j]
                purpleindex.append(coord)        
            elif np.average(abs(np.subtract(red,img[j,i,:]))) < sensitivity:
                coord = [i,j]
                redindex.append(coord)    
            else:
                pass
                  
    # Interpolation of values
    # Identify location of xmin, xmax, ymin and ymax
    
    xmin= 89
    xmax = 910
    ymin = 571
    ymax = 87
    
    xpixel = xmax-xmin
    ypixel = ymin-ymax
    
    xrange = [0,350]
    yrange = [4,-6]
    
    xlin = np.linspace(xrange[0],xrange[1],xpixel)
    ylin = np.linspace(yrange[0],yrange[1],ypixel)
            
    greeninterp = []
    blueinterp = []
    purpleinterp = [] 
    redinterp = []  
            
    for i in range(len(greenindex)):
        greenx = xlin[greenindex[i][0]-xmin]
        greeny = ylin[greenindex[i][1]-ymax]
        gg = [greenx,greeny]
        greeninterp.append(gg)
    
    for i in range(len(blueindex)):
        bluex = xlin[blueindex[i][0]-xmin]
        bluey = ylin[blueindex[i][1]-ymax]
        bb = [bluex,bluey]
        blueinterp.append(bb)
        
    for i in range(len(purpleindex)):
        purplex = xlin[purpleindex[i][0]-xmin]
        purpley = ylin[purpleindex[i][1]-ymax]
        pp = [purplex,purpley]
        purpleinterp.append(pp)
        
    for i in range(len(redindex)):
        redx = xlin[redindex[i][0]-xmin]
        redy = ylin[redindex[i][1]-ymax]
        rr = [redx,redy]
        redinterp.append(rr)
    
    gex = pd.DataFrame(greeninterp)
    bex = pd.DataFrame(blueinterp)
    pex = pd.DataFrame(purpleinterp)
    rex = pd.DataFrame(redinterp)
    
    
    compiled = pd.concat([gex,bex,pex,rex], axis = 1)
    compiled.columns = ['gx','gy','bx','by','px','py','rx','ry']
    
    compiled.to_csv(csvName,index = False)
    
    
t = time.time()

for i in range(1):
    for j in range(8):
        picName = 'WindTapSourcePage' + str(i+1) + ',' + str(j) + '.jpg'
        docName = 'csvWindTapSourcePage' + str(i+1) + ',' + str(j) + '.csv'
        jpgtoCsv(picName, docName)

elapsed = time.time()-t
print(str(elapsed) + 'sec')
