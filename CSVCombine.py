# -*- coding: utf-8 -*-
"""
Created on Fri Jul 16 16:56:32 2021

@author: brendanlia
"""
import os
import glob
import pandas as pd
os.chdir(r"C:\Users\brendanlia\Desktop\Jacking Tower\CSVs")

extension = 'csv'
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]

#combine all files in the list
combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames ])
#export to csv
combined_csv.to_csv( "combined_csv.csv", index=False, encoding='utf-8-sig')

