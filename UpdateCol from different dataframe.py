#Unzip the file
from zipfile import ZipFile
zf = ZipFile(r"C:\xxx\Schema_xxx.zip")
zf.extractall(r"C:\xxx\Task xxx")
zf.close()
#Import Excel and TXT Files
import pandas as pd
import numpy as np
dfexcel = pd.read_excel (r'C:xxxx\xyz.xlsx', usecols=['UKPRN','Number'], dtype={'UKPRN':str,'Number':str}) 
dftxt = pd.read_csv(r'C:\xxx\zyx.txt', delimiter="\t", dtype={'id':str,'__l9extN':str})
#Rename col to id in the excel dataset
dfexcel.rename(columns = {'UKPRN':'id', 'Number':'__l9extN'}, inplace=True)
#set index to id
dfexcel.set_index('id', inplace=True)
dftxt.set_index('id', inplace=True)
#update the txt df with excel df
dftxt.update(dfexcel)
#Export the file back to txt
dftxt.to_csv(r'C:\xxxx\Schema_xxx.txt', sep='\t')
