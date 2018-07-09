import pandas as pd     #importing Pandas Package to wrangle data
import glob 
import os 
import re

### List Source Files That I need to Import###
path  = os.getcwd()
files = os.listdir(path)

### Loading Files by Variable ###

df   = pd.DataFrame()
data = pd.DataFrame()


for files in glob.glob('O:/My Drive/Colgate_Australia/Analytics_Development/Data_Governance/Master_Data/Sales_Master_Data/Aldi_Weekly_Sales/ALDI Supplier *.xlsx'):
    
    data = pd.read_excel(files,'Sheet1',skiprows=1).fillna(method='ffill')
    date = re.compile(r'([\.\d]+ - [\.\d]+)').search(files).groups()[0] # assigns Date within file name to df
    data['Date'] = date
    data['Start_Date'], data['End_Date'] = data['Date'].str.split(' - ', 1).str
    data['End_Date'] = data['End_Date'].str[:10]
    data['Start_Date'] = data['Start_Date'].str[:10]
    data['Start_Date'] =pd.to_datetime(data['Start_Date'],format ='%d.%m.%Y',errors='coerce') 
    data['End_Date']= pd.to_datetime(data['End_Date'],format ='%d.%m.%Y',errors='coerce')
    df  = df.append(data)
    df
    
df.drop(['Date'], axis=1)  

#Linking Pricing from another file
df2  = pd.DataFrame()
file1 = 'O:/My Drive/Colgate_Australia/Analytics_Development/Data_Governance/Master_Data/Sales_Master_Data/Aldi_Weekly_Sales/Financial Information.xlsx'
df2 = pd.read_excel(file1,'Sheet1')


dfmerge= df.merge(df2, how='left', left_on=['Code'], right_on=['Code'])  

dfmerge['Value'] = dfmerge['Sold (Units)']*dfmerge['RRP']

dfmerge = dfmerge.sort_values(['Start_Date'], 
                                      ascending= True, kind ='quicksort',na_position = 'last')

dfmerge.to_excel('OUTPUT ALDI Data File.xlsx',index= True,merge_cells=False,startrow=1)
