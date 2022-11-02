#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import numpy as np
import glob
from pathlib import Path
from datetime import datetime as dt


# In[2]:


home = Path.home()
today = dt.today()
today = today.strftime("%Y%m%d")
today = '20221101'


# In[3]:


target = Path (home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Single shortage')
PNFVPath = Path(home, 'HP Inc','GPSTW SOP - 2021 日新', 'PN FV description mapping table_ALL.xlsx')
ExternalReportFolder = Path(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination')
ExternalReport = [f for f in glob.glob(str(Path(ExternalReportFolder, today + '*')))]
ExternalReport


# In[4]:


PNFVFile = pd.read_excel(PNFVPath)
PNFVFile = PNFVFile [['PN', 'Descr']]


# In[5]:


ODMdict = {
    'FWH' : 'WHFXN',
    'Compal' : 'KSCEI',
    'CEI' : 'KSCEI',
    'Wistron' : 'CQWIS',
    'Inventec' : 'CQIEC',
    'Quanta' : 'CQQCI',
    'Pegatron' : 'CQPCQ'
}


# In[6]:


fileList = [str(x) for x in target.glob("*xlsx")]


# In[7]:


errorList = []


# In[8]:


resultList = []


# In[9]:



def clean(fname: str, file : pd.DataFrame) -> pd.DataFrame:

    #add report day
    currentYear = dt.now().year
    #print(fname)
    currentday = fname.split('\\')[-1][-13:-5]
    #currentday = str(currentYear) + currentday
    #print(currentday)
    #print(type(currentday))
    file = file.assign(reportDate = currentday)
    file['reportDate'] = file['reportDate'].apply(lambda x: dt.strptime(x, '%Y%m%d'))
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


# In[10]:


for f in fileList:
    # file = pd.read_excel(f)
    # resultList.append(clean(f, file))


    try:
        file = pd.read_excel(f)
        resultList.append(clean(f, file))
        print(f + " process done!")
    except Exception as e:
        errorList.append([f, e])
        print(f + " process failed!")


# In[11]:


result = pd.concat(resultList)


# In[12]:


dateList = result['reportDate'].tolist()


# In[13]:


max(dateList)


# In[14]:


LatestSGMaterial = result[result['reportDate'] == max(dateList)]


# In[15]:


LatestSGMaterial


# result['Prev_Single Shortage QTY'].unique()

# In[16]:


LatestSGMaterial = LatestSGMaterial.merge(PNFVFile.rename(columns = {'PN': 'HP PN'}), on = 'HP PN', how = 'left')
LatestSGMaterial


# In[17]:


LatestSGMaterial['Key'] = LatestSGMaterial['ODM'] + LatestSGMaterial['Descr']
KeyList = LatestSGMaterial['Key'].tolist()


# In[18]:


ExternalReport


# In[19]:


externalResultDFList = []


# In[20]:


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


# In[21]:


externalResultDF = pd.concat(externalResultDFList)


# In[22]:


#externalResultDF['test'] = externalResultDF['GPS Remark'].apply(lambda x : str(x).split("\n"))


# In[23]:


#externalResultDF['test'] = externalResultDF['test'].apply(lambda x: i.replace() for i in x)


# In[24]:


PNFVFile


# In[25]:


PNFVFile = PNFVFile.rename(columns = {'PN': 'HP PN'})


# In[26]:


result = result.merge(PNFVFile.rename(columns = {'PN': 'HP PN'}), on = 'HP PN', how = 'left')


# In[27]:


result.head()


# In[28]:


result = result.merge(externalResultDF.rename(columns = {'FV/Des' : 'Descr'}), on = ['ODM', 'Descr'], how = 'left')


# In[29]:


result


# In[30]:


result.to_excel(Path(target, 'total singal shortage_' + today +'.xlsx'), index = False)


# In[31]:


errorList


# In[32]:


import psutil


# In[33]:


psutil.cpu_percent()


# In[34]:


psutil.virtual_memory()


# In[35]:


print(psutil.Process(os.getpid()).memory_info().rss / 1024 ** 2)


# In[ ]:




