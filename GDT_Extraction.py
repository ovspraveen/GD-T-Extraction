#!/usr/bin/env python
import pandas as pd 
import numpy as np
import openpyxl as px
wb = px.lo ad_workbook('name_of_excel_sheet.xlsx')
num_sheet=len(wb.sheetnames)
df_full={}
i=5
for i in range(num_sheet):
    df=pd.read_excel('name_of_excel_sheet.xlsx',sheet_name=i)
    head=df.iloc[1]
    df_full[i]=df[2:]
    df_full[i].columns=head
i=0
for i in range(5):
    df_full[i]=pd.read_excel('name_of_excel_sheet.xlsx',sheet_name=i)

df=df_full[2]
sheet=wb.get_sheet_by_name(wb.sheetnames[2])
for i in range(len(df)):
    if sheet['A'+str(i+1)].value=='ID':
        k=i
        break
        
head=df.iloc[k-1]
df=df[k:]
df.columns=head
a=df
for i in (a.index):
    if 'nan' in (str(a['Entity'][i])):
        break
a=df[:i]
a.dropna(axis=1,inplace=True)      
df_full[2]=a
df_full[2]['ID']=df_full[2].index
df_full[2].index=range(1,len(df_full[2])+1)

df_full[2]['Type']=float('nan')
for i in range(len(df_full[2])):
    if 'datum_system' in df_full[2]['Entity'][i+1]:
        df_full[2]['Type'][i+1]='datum'
    elif 'dimensional_characteristic_representation' in df_full[2]['Entity'][i+1]:
        df_full[2]['Type'][i+1]='dimension'
    elif 'tolerance' in df_full[2]['Entity'][i+1]:
        df_full[2]['Type'][i+1]='tolerance'
df_full[2]['Dimension']=0
df_full[2]['Tolerance']=0
for i in range(len(df_full[2])):
    if df_full[2]['Type'][i+1]=='dimension':
        if (('∅' in df_full[2]['PMI Representation'][i+1])==True) & (('±' in df_full[2]['PMI Representation'][i+1])==False):
            df_full[2]['Dimension'][i+1]=float(df_full[2]['PMI Representation'][i+1].split('∅')[1])
        elif (('∅' in df_full[2]['PMI Representation'][i+1])==True) & (('±' in df_full[2]['PMI Representation'][i+1])==True):
            df_full[2]['Dimension'][i+1]=float(df_full[2]['PMI Representation'][i+1].split('∅')[1].split('±')[0])
            df_full[2]['Tolerance'][i+1]=(df_full[2]['PMI Representation'][i+1].split('∅')[1].split('±')[1])
        elif (('∅' in df_full[2]['PMI Representation'][i+1])==False) & (('±' in df_full[2]['PMI Representation'][i+1])==True):
            df_full[2]['Dimension'][i+1]=float(df_full[2]['PMI Representation'][i+1].split('±')[0])    
            df_full[2]['Tolerance'][i+1]=(df_full[2]['PMI Representation'][i+1].split('±')[1])
        elif (('±' in df_full[2]['PMI Representation'][i+1])==False) & (('∅' in df_full[2]['PMI Representation'][i+1])==False):
            df_full[2]['Dimension'][i+1]=float(df_full[2]['PMI Representation'][i+1])

data=[list(df_full[2]['Dimension']),list(df_full[2]['Tolerance'])]
values=pd.DataFrame(data)
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
values.to_excel(writer, sheet_name='test_1', index=False)
writer.close()
