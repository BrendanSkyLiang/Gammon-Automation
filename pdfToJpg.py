# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 13:49:55 2021

@author: brendanlia
"""

import cv2 as cv
import numpy as np
import glob
import matplotlib.pyplot as plt
import pandas as pd
import plotdigitizer


for i in range(213):
    pagenum = str(i+1)
    docname = 'WindTapSource Page'+ pagenum + '.jpg'
    page = cv.imread(docname)
   
    Crop11 = page[72:730,62:1024]
    num = str(0)
    name11 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name11,Crop11)
    
    Crop12 = page[72:730,1039:2001]
    num = str(1)
    name12 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name12,Crop12)
    
    Crop21 = page[752:1410,62:1024]
    num = str(2)
    name21 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name21,Crop21)
    
    Crop22 = page[752:1410,1039:2001]
    num = str(3)
    name22 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name22,Crop22)
    
    Crop31 = page[1432:2090,62:1024]
    num = str(4)
    name31 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name31,Crop31)
    
    Crop32 = page[1432:2090,1039:2001]
    num = str(5)
    name32 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name32,Crop32)
    
    Crop41 = page[2114:2772,62:1024]
    num = str(6)
    name41 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name41,Crop41)
    
    Crop42 = page[2114:2772,1039:2001]
    num = str(7)
    name42 = 'WindTapSourcePage' + pagenum + ',' + num + '.jpg'
    cv.imwrite(name42,Crop42)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    