# -*- coding: utf-8 -*-
"""
Created on Thu Jul 29 13:02:54 2021

@author: brendanlia
"""

from cv2 import imread
from numpy import average
from numpy import subtract
from numpy import linspace
import pandas as pd
import time

# cv uses BGR while matplotlib uses RGB
# Load image as RGB 3D array

def jpgtoCsv(jpgName,csvName):
    img = imread(jpgName)
    # img = cv.cvtColor(img, cv.COLOR_BGR2RGB)
    # # plt.imshow(img)
    
    # # Find indexes of RGB values close to the line colours
    sensitivity = 5 # <----- can be changed
    # green = [158,171,117]
    # blue = [104,125,162]
    # purple = [126,113,140]
    # red = [150,91,94]
    
    green = [117,171,158]
    blue = [162,125,104]
    purple = [140,113,126]
    red = [94,91,150]
    
    greenindex = []
    blueindex = []
    purpleindex = []
    redindex = []
    
    for i in range(len(img[1])):
        for j in range(len(img)):
            if img[j,i,0] > 180:
                pass
            elif img[j,i,0] < 70:
                pass
            elif img[j,i,2] < 80:
                pass
            elif average(abs(subtract(green,img[j,i,:]))) < sensitivity:
                greenindex.append([i,j])
            elif average(abs(subtract(blue,img[j,i,:]))) < sensitivity:
                blueindex.append([i,j])
            elif average(abs(subtract(purple,img[j,i,:]))) < sensitivity:
                purpleindex.append([i,j])        
            elif average(abs(subtract(red,img[j,i,:]))) < sensitivity:
                redindex.append([i,j])    
            else:
                pass
                  
    # Interpolation of values
    # Identify location of xmin, xmax, ymin and ymax
    xmin, xmax, ymin, ymax = 89, 910, 571, 87
    # xmax = 910
    # ymin = 571
    # ymax = 87
    xpixel = xmax-xmin
    ypixel = ymin-ymax
    xrange = [0,350]
    yrange = [4,-6]
    
    xlin = linspace(xrange[0],xrange[1],xpixel)
    ylin = linspace(yrange[0],yrange[1],ypixel)
            
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

for i in range(1):  # <--- number of pages
    for j in range(8): # <--- number of graphs per page
        picName = 'WindTapSourcePage' + str(i+1) + ',' + str(j) + '.jpg'
        docName = 'csvWindTapSourcePage' + str(i+1) + ',' + str(j) + '.csv'
        jpgtoCsv(picName, docName)
        
elapsed = time.time()-t
print(str(elapsed) + 'sec')
