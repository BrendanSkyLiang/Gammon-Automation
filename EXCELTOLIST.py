# -*- coding: utf-8 -*-
"""
Created on Mon Jul 19 11:51:20 2021

@author: brendanlia
"""

import pandas

def flatten(t):
    return [item for sublist in t for item in sublist]

test = pandas.read_clipboard(header=None)  
array = test.values.tolist()
Ax = flatten(array)
Ex = []

for item in Ax:
    Ex.append(str(item))

while 'nan' in Ex: Ex.remove('nan')
Ex1 = [x.replace('.0','') for x in Ex]




































