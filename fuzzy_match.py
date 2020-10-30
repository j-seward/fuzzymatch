#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 17:35:13 2020

@author: Jonathan Seward

Program: This program allows fuzzy matching from strings in a Stata dataset to 
    an excel file. It was based on an online tutorial, which I can no longer 
    find so at least some of the commands are not my creation. I used Florida's
    AHCA data and the SK&A dataset to match hospital names, but this should be 
    adaptable to multiple datasets.
"""

# set read file
infile = ''
#set lookup file
lookupfile = '' 
#set output file
outfile = ''

# options for matching: fuzz.ratio, fuzz.partial_ratio, fuzz.token_sort_ratio, fuzz.token_set_ratio
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def checker(wrong_options,correct_options):
    names_array=[]
    ratio_array=[]    
    for wrong_option in wrong_options:
        if wrong_option in correct_options:
           names_array.append(wrong_option)
           ratio_array.append(100)
        else:   
            x=process.extractOne(wrong_option,correct_options,scorer=fuzz.ratio)
            names_array.append(x[0])
            ratio_array.append(x[1])
    return names_array,ratio_array

orig_list = pd.read_stata(infile)
lookup_table = pd.read_excel(lookupfile)

# fixes names on lookuptable; column names are dependent on dataset
lookup_table = lookup_table[['hospital_name','health_sys_name','hospital_city']]
for col in lookup_table.columns: 
    print(col)
    
# add original name to match name so that you can look it up
str2Match = orig_list['hospital_name'].fillna('######').tolist()
strOptions =lookup_table['hospital_name'].fillna('######').tolist()

# run the matching function
name_match,ratio_match=checker(str2Match,strOptions)

# put everything into a dataframe; column names are dependent on dataset
df1 = pd.DataFrame()
df1['faclnbr']=orig_list['faclnbr']
df1['orig_names']=pd.Series(str2Match)
df1['matched_names']=pd.Series(name_match)
df1['correct_ratio']=pd.Series(ratio_match)

# need to also grab health_sys_name
df1 = df1.merge(lookup_table,left_on='matched_names',right_on='hospital_name')


# Get names of indexes for which column match is greater than 0
indexNames = df1[df1['correct_ratio'] < 0 ].index
 
# Delete these row indexes from dataFrame
df1.drop(indexNames , inplace=True)

# Output to excel
df1.to_excel(outfile, engine='xlsxwriter')
