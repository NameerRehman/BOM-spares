# -*- coding: utf-8 -*-
"""
Created on Tue Aug 18 16:08:28 2020

@author: nrehman
"""

import pandas as pd
import glob, os
from xlrd import open_workbook

#Read all xls files in Modules folder
path = "C:/Users/nrehman/Documents/Modules"
all_modules = glob.glob(path + "/*.xls")

#Append each file contents to dataframe
bom = pd.concat((pd.read_excel(m, header=7).assign(MODULE=os.path.basename(m)) for m in all_modules))

#drop all rows containing NAs in Spare Class columns
bom = bom.dropna(subset = ['SPARE CLASS'])

#Convert to string
bom['SPARE CLASS'] = bom['SPARE CLASS'].astype(str)
bom['MODULE'] = bom['MODULE'].astype(str)

#Strip Spaces
bom['SPARE CLASS'] = bom['SPARE CLASS'].str.replace(' ','')
bom['MODULE'] = bom['MODULE'].str.replace('.xls','')

#Sort df by Spare Class values
bom['SPARE CLASS'] = pd.Categorical(bom['SPARE CLASS'], ['1','2','3','9','X'])
bom = bom.sort_values('SPARE CLASS')

#Create new df without Space Class X items
bom_spares = bom[bom['SPARE CLASS'] != 'X']

#Create new df without duplicates
bom_spares_unique = bom_spares.drop_duplicates(subset = ['PART NUMBER'])

#Iterate through each unique part in df
for i in range(bom_spares_unique.shape[0]):
    part = bom_spares_unique['PART NUMBER'].iloc[i]
    
    #filter bom_spares by "part" & output MODULE column to a list
    modules_list = bom_spares[bom_spares['PART NUMBER'] == part]['MODULE'].to_list()
    
    #assign all module occurences of "part" to bom_spares_unique MODULE column
    bom_spares_unique['MODULE'].iloc[i] = modules_list
    
    #filter bom_spares by "part" & output PROJ QTY column to a list
    total_qty = bom_spares[bom_spares['PART NUMBER'] == part]['PROJ\nQTY.'].to_list()
    
    #Calculate sum of list
    total_qty_sum = 0
    for j in total_qty:
        total_qty_sum+=j
    bom_spares_unique['PROJ\nQTY.'].iloc[i] = total_qty_sum

#Export df to excel sheet
with pd.ExcelWriter('Spare Parts List.xlsx') as writer:
    bom.to_excel(writer, sheet_name='All Parts')
    bom_spares.to_excel(writer, sheet_name='Spares')
    bom_spares_unique.to_excel(writer, sheet_name='Spares Unqiue')


