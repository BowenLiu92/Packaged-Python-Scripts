#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd

df = pd.read_excel(r'file:///L:\lvpcgis\BowenLiuTemporary\I-78%20Adams%20Rd%20Point%20of%20Access%20Study/ExcelFiles\Aug%20Updated%20Results\New%20Results.xlsx', sheet_name = input('Sheet Name'))
df.columns =['TAZ', 'MUNICODE', 'NAME', 'COUNTY', 2010, 2015,
       2020, 2025, 2030, 2035,
       2040, 2045]
df['New2015'] = df[2015]

#The function that we are going to use
def letstry(df, n, growthsheet, county):
    '''df is the data frame with reference data, in our case the old population distribution data
    n is the year number, growthsheet is the dataframe that contains the population increase data,
    county is either 'LC' or 'NC' '''
    df[str(n-5)+ '-' +str(n)] = df[n] - df[n-5]
    growth = df[str(n-5)+ '-' +str(n)].sum()
    df[str(n-5)+ '-' +str(n)+ 'ratio']  = df[str(n-5)+ '-' +str(n)]/growth
    realg = growthsheet.at[county, n]
    df['New' + str(n)] = df['New' + str(n-5)] + realg*df[str(n-5)+ '-' +str(n)+ 'ratio']

#Do LC first
df1 = df[df['COUNTY'] == 77]
growthsheet = pd.read_excel(input('Newest growth sheet file location'))

for i in range(2020, 2050, 5):
    letstry(df1, i, growthsheet, 'LC')

collist = [ '2015-2020ratio',        'New2020',      '2020-2025',
   '2020-2025ratio',        'New2025',      '2025-2030', '2025-2030ratio',
          'New2030',      '2030-2035', '2030-2035ratio',        'New2035',
        '2035-2040', '2035-2040ratio',        'New2040',      '2040-2045',
   '2040-2045ratio',        'New2045']
df1.loc['sum'] = df1[collist].sum()

#Do NC next
df2 = df[df['COUNTY'] == 95]

for i in range(2020, 2050, 5):
    letstry(df2, i, growthsheet, 'NC')

collist = [ '2015-2020ratio',        'New2020',      '2020-2025',
   '2020-2025ratio',        'New2025',      '2025-2030', '2025-2030ratio',
          'New2030',      '2030-2035', '2030-2035ratio',        'New2035',
        '2035-2040', '2035-2040ratio',        'New2040',      '2040-2045',
   '2040-2045ratio',        'New2045']
df2.loc['sum'] = df2[collist].sum()

#Combine LC and NC
#Drop the sum row for df1 and df2
df3_1 = df1.drop(['sum'])
df3_2 = df2.drop(['sum'])

df3 = pd.concat([df3_1, df3_2]).sort_values('TAZ')


df1 = df1[[ 'TAZ',       'MUNICODE',           'NAME',         'COUNTY', 'New2015',
         'New2020', 'New2025', 'New2030', 'New2035', 'New2040', 'New2045']]
df2 = df2[[ 'TAZ',       'MUNICODE',           'NAME',         'COUNTY', 'New2015',
         'New2020', 'New2025', 'New2030', 'New2035', 'New2040', 'New2045']]
df3 = df3[[ 'TAZ',       'MUNICODE',           'NAME',         'COUNTY', 'New2015',
         'New2020', 'New2025', 'New2030', 'New2035', 'New2040', 'New2045']]

folderfilename = input('Give me the file location and also the file name')


writer = pd.ExcelWriter(folderfilename)
df1.to_excel(writer,'LC')
df2.to_excel(writer,'NC')
df3.to_excel(writer,'LV')
writer.save()

