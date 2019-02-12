#!/usr/bin/env python
# coding: utf-8

# In[1]:


from IPython.core.interactiveshell import InteractiveShell
InteractiveShell.ast_node_interactivity = "all"

import numpy as np
import pandas as pd

def reportPE(filename, sheetname):
    df = pd.read_excel(filename, sheet_name = sheetname, header = 2)
    df = df.iloc[:, np.r_[1:5, 5:len(df.columns):2]]
    df = df.groupby(['NAME', 'COUNTY'])['New 2010 EMP', 'New 2015 EMP',
       'New 2020 EMP', 'New 2025 EMP', 'New 2030 EMP', 'New 2035 EMP',
       'New 2040 EMP', 'New 2045 EMP'].sum().reset_index()
    df = df.append(df.sum(numeric_only=True), ignore_index=True)
    return df

LC_filename = 'file:///\\Fileserver01\d_drive\lvpcprojects\Pop%20+%20Emp%20Projections%202018/Feb%202019%20Update\Employment%20Growth%20Rate%20and%20Update.xlsx'
LC_sheetname = 'LC EMP Update'

NC_filename = 'file:///\\Fileserver01\d_drive\lvpcprojects\Pop%20+%20Emp%20Projections%202018/Feb%202019%20Update\Employment%20Growth%20Rate%20and%20Update.xlsx'
NC_sheetname = 'NC EMP Update'

LC_EMP = reportPE(LC_filename, LC_sheetname)
NC_EMP = reportPE(NC_filename, NC_sheetname)
writer = pd.ExcelWriter('H:\lvpcprojects\Pop + Emp Projections 2018\Feb 2019 Update\Results\EMP_output.xlsx')
LC_EMP.to_excel(writer,'LC EMP')
NC_EMP.to_excel(writer,'NC EMP')

writer.save()

