import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

data=pd.read_excel('superstore_dataset2011-2015.xls',sheet_name = 'superstore_dataset2011-2015' )
data.head()

data.shape

data.dtypes

data1=data.loc[0:20066,:]
data2=data.loc[20067:51289,:]

data1.loc[:,'Order Date']=pd.to_datetime(data.loc[:,'Order Date'],format='%d/%m/%Y',errors='coerce')

data2.loc[:,'Order Date']=pd.to_datetime(data.loc[:,'Order Date'],format='%d-%m-%Y',errors='coerce')

data=data1.append(data2)

