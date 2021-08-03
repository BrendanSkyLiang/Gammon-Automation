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
from time import time

# cv uses BGR while matplotlib uses RGB
# Load image as RGB 3D array

def jpgtoCsv(jpgName,csvName):
    img = imread(jpgName)
    
    # Find indexes of RGB values close to the line colours
    sensitivity = 5 # <----- can be changed
    
    
    # col  = [B,G,R]
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
                  
    # Interpolate values
    # Identify location of xmin, xmax, ymin and ymax
    xmin, xmax, ymin, ymax = 89, 910, 573, 87
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
        greeninterp.append([greenx,greeny])
    
    for i in range(len(blueindex)):
        bluex = xlin[blueindex[i][0]-xmin]
        bluey = ylin[blueindex[i][1]-ymax]
        blueinterp.append([bluex,bluey])
        
    for i in range(len(purpleindex)):
        purplex = xlin[purpleindex[i][0]-xmin]
        purpley = ylin[purpleindex[i][1]-ymax]
        purpleinterp.append([purplex,purpley])
        
    for i in range(len(redindex)):
        redx = xlin[redindex[i][0]-xmin]
        redy = ylin[redindex[i][1]-ymax]
        redinterp.append([redx,redy])
    
    gex = pd.DataFrame(greeninterp)
    bex = pd.DataFrame(blueinterp)
    pex = pd.DataFrame(purpleinterp)
    rex = pd.DataFrame(redinterp)
    
    compiled = pd.concat([gex,bex,pex,rex], axis = 1)
    compiled.columns = ['gx','gy','bx','by','px','py','rx','ry']
    
    compiled.to_csv(csvName,index = False)

t = time()

for i in range(180,213):  # <--- number of pages Note, range(a,b) = a, a+1, ... , b-1:  it doesn't include b
    for j in range(8): # <--- number of graphs per page
        picName = 'WindTapSourcePage' + str(i+1) + ',' + str(j) + '.jpg'
        docName = 'Pg' + str(i+1) + ',' + str(j) + '.csv'
        print(docName)
        jpgtoCsv(picName, docName)

elapsed = time()-t
print(str(elapsed) + 'sec')
