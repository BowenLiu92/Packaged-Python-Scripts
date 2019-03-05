#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#This .py file is created to automate the yearly project of distributing employment across the
# 4 categories: Office, Industry, Retail, Other

import pandas as pd
import numpy as np

colist = ['TAZ', 'MUNICODE', 'NAME', 'COUNTY', 'New2015', 'New2020', 'New2025',
       'New2030', 'New2035', 'New2040', 'New2045']
df = pd.read_excel('file:///H:\lvpcprojects\Pop%20+%20Emp%20Projections%202018\Feb%202019%20Update\TAZ%20based%20results\EMP%20result.xlsx',
                    sheet_name = 'LV')[colist]
distri_file = r'file:///L:\lvpcgis\BowenLiuTemporary\I-78%20Adams%20Rd%20Point%20of%20Access%20Study/ExcelFiles\Aug%20Updated%20Results\New%20Results.xlsx'
def distribute_emp(sheet, year):
    df1 = pd.read_excel(distri_file, sheet_name = sheet) 
    df1['Retail Ratio ' + year] = df1['Retail']/df1['Bowen ' + year + ' Total']
    df1['Office Ratio ' + year] = df1['Office']/df1['Bowen ' + year + ' Total']
    df1['INDUST Ratio ' + year] = df1['INDUST ']/df1['Bowen ' + year + ' Total']
    df1['OTHER Ratio ' + year] = 1 - df1[['Retail Ratio ' + year, 'Office Ratio ' + year, 'INDUST Ratio ' + year]].sum(axis = 1)
    
    #Finish the ratio, now I need to calculate the new distribution number based on the ratio.
    df1['New Retail '+ year] = df['New'+year] * df1['Retail Ratio ' + year]
    df1['New Office '+ year] = df['New'+year] * df1['Office Ratio ' + year]
    df1['New INDUST '+ year] = df['New'+year] * df1['INDUST Ratio ' + year]
    df1['New OTHER '+ year] = df['New'+year] * df1['OTHER Ratio ' + year]
    dfm = pd.merge(df, df1.set_index('CODE'), left_on= 'TAZ', right_index= True)
    
    #Filter the columns. After all I only need to see the columns showing the results.
    all_columns = list(dfm.columns)
    clist = [e for e in all_columns if e in ['TAZ', 'MUNICODE', 'NAME', 'New Retail '+year, 'New Office '+year,
       'New INDUST '+year, 'New OTHER '+year]]
    dfm.fillna(0, inplace = True)
    dfm = dfm[clist]
    
    return dfm

d = {}
for i in range(2020, 2050, 5):
    d[i] = distribute_emp(sheet = str(i) + ' Distri', year = str(i))
    
writer = pd.ExcelWriter(input('File location and file name: '))
for e in d.keys():
    d[e].to_excel(writer, sheet_name = 'Distribution '+ str(e))
writer.save()

