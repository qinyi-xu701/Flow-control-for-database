#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import numpy as np
import glob
from pathlib import Path
from datetime import datetime as dt
from datetime import timedelta


# In[4]:


home = Path.home()
today = dt.today()

dateRange = [today - timedelta(days = x) for x in range(100)]
dateRange = [i.strftime("%Y%m%d") for i in dateRange]

today = today.strftime("%Y%m%d")
today = '20221108'

ODMdict = {
    'FWH' : 'WHFXN',
    'Compal' : 'KSCEI',
    'CEI' : 'KSCEI',
    'Wistron' : 'CQWIS',
    'Inventec' : 'CQIEC',
    'Quanta' : 'CQQCI',
    'Pegatron' : 'CQPCQ'
}


# In[5]:


def clean(fname: str, file : pd.DataFrame) -> pd.DataFrame:

    currentYear = dt.now().year
    currentday = fname.split('\\')[-1][-13:-5]
    file = file.assign(LastSGreportDate = currentday)
    
    file['LastSGreportDate'] = file['LastSGreportDate'].apply(lambda x: dt.strptime(x, '%Y%m%d'))
    file['LastSGreportDate'] = pd.to_datetime(file['LastSGreportDate'])

    file = file.assign(reportDate = today)
    file['reportDate'] = pd.to_datetime(file['reportDate'])

    #clean
    file.columns = file.columns.str.strip()

    #drop useless columns and rows
    file = file.drop(columns = ['Description (Item)', 'Schedule (Comments)', 'Hub inventory', 'Vendor'])
    file = file[file['Procurement type'] == 'B/S'].reset_index(drop = True)

    #adjust qty columns name and units
    qtycol = file.filter(like='Single Shortage QTY (K)').columns.tolist()
    

    for i in qtycol:
        file[i] = file[i].apply(lambda x: x.upper() if type(x) == str else x)
        file[i] = file[i].replace("NEW", 0)
        file[i] = file[i].apply(lambda x: x*1000)
    file = file.rename(columns= {qtycol[0]: 'Single Shortage QTY', qtycol[1]: 'Prev_Single Shortage QTY'})

    #replace ODM name
    file['ODM'] = file['ODM'].replace(ODMdict)
    return file   


# In[6]:


target = Path (home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Single shortage')
PNFVPath = Path(home, 'HP Inc','GPSTW SOP - 2021 日新', 'PN FV description mapping table_ALL.xlsx')



PNFVFile = pd.read_excel(PNFVPath)
PNFVFile = PNFVFile [['PN', 'Descr']]
PNFVFile = PNFVFile.rename(columns = {'PN': 'HP PN'})


# In[7]:


fileList = [str(x) for x in target.glob("*xlsx")]
errorList = []
resultList = []


# In[8]:


for f in fileList:
    try:
        file = pd.read_excel(f)
        resultList.append(clean(f, file))
        print(f + " process done!")
    except Exception as e:
        errorList.append([f, e])
        print(f + " process failed!")


# In[9]:


result = pd.concat(resultList)


# In[10]:


dateList = result['reportDate'].tolist()
max(dateList)


# In[ ]:


LatestSGMaterial = result
LatestSGMaterial = LatestSGMaterial.merge(PNFVFile, on = 'HP PN', how = 'left')
LatestSGMaterial['Key'] = LatestSGMaterial['ODM'] + LatestSGMaterial['Descr']
KeyList = LatestSGMaterial['Key'].tolist()


# ### concat current day external report

# In[ ]:


ExternalReportFolder = Path(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination')
ExternalReport = [f for f in glob.glob(str(Path(ExternalReportFolder, today + '*')))]
externalResultDFList = []

for _ in ExternalReport:
    try: 
        temp = pd.read_excel(_)
        temp['ODM'] = temp['ODM'].ffill()
        temp['ODM'] = temp['ODM'].replace(ODMdict)
        temp['FV/Des'] = temp['FV/Des'].ffill()
        #temp['ETA'] = temp['ETA'].ffill()
        temp['key'] = temp['ODM'] + temp['FV/Des']
        temp = temp[temp.key.isin(KeyList)]
        temp = temp[['ODM', 'FV/Des', 'HP_PN', 'ETA', 'GPS Remark']]
        temp = temp.groupby(['ODM', 'FV/Des']).agg({'ETA' : lambda x: '\n'.join(set(x.dropna())),
                                                    'GPS Remark': lambda x: '\n'.join(set(x.dropna()))})
        temp = temp.reset_index()
        if len(temp) > 0:
            print(len(temp))
            externalResultDFList.append(temp)
        else:
            pass
    except Exception as e:
        print(e)
        print(_)

externalResultDF = pd.concat(externalResultDFList)


# ### lookup PNFV and merge external reportm

# In[ ]:


result = result.merge(PNFVFile.rename(columns = {'PN': 'HP PN'}), on = 'HP PN', how = 'left')
result = result.merge(externalResultDF.rename(columns = {'FV/Des' : 'Descr'}), on = ['ODM', 'Descr'], how = 'left')
result = result.drop_duplicates()


# ### output

# In[ ]:


result.to_excel(Path(target, 'total singal shortage_' + today +'.xls'), index = False)


# In[ ]:


errorList


# In[ ]:


import psutil
psutil.cpu_percent()
psutil.virtual_memory()
print(psutil.Process(os.getpid()).memory_info().rss / 1024 ** 2)


# In[ ]:




