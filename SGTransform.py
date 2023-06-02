#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import os
import numpy as np
import glob
from pathlib import Path
from datetime import datetime as dt
from datetime import timedelta
import pyodbc
import shutil


# In[ ]:


home = Path.home()
today = dt.today().date()

dateRange = [today - timedelta(days = x) for x in range(365)]
#dateRange = [i.strftime("%Y%m%d") for i in dateRange]

today = today.strftime("%Y%m%d")


ODMdict = {
    'FWH' : 'WHFXN',
    'FXN' : 'WHFXN',
    'Foxconn WH' : 'WHFXN',
    'Compal' : 'KSCEI',
    'CEI' : 'KSCEI',
    'Wistron' : 'CQWIS',
    'wistron' : 'CQWIS',
    'WCQ': 'CQWIS',
    'Inventec' : 'CQIEC',
    'Inventec CQ' : 'CQIEC',
    'Quanta' : 'CQQCI',
    'Pegatron' : 'CQPCQ'
}


# In[ ]:


dateRange.reverse()


# In[ ]:


def clean(fname: str, file : pd.DataFrame, externalReportDate : str) -> pd.DataFrame:

    currentYear = dt.now().year
    currentday = fname.split('\\')[-1][-13:-5]
    file = file.assign(LastSGreportDate = currentday)
    
    file['LastSGreportDate'] = file['LastSGreportDate'].apply(lambda x: dt.strptime(x, '%Y%m%d'))
    file['LastSGreportDate'] = pd.to_datetime(file['LastSGreportDate'])

    file = file.assign(reportDate = externalReportDate)

    file['reportDate'] = file['reportDate'].apply(lambda x: dt.strptime(x, '%Y%m%d'))
    file['reportDate'] = pd.to_datetime(file['reportDate'])

    #clean
    file.columns = file.columns.str.strip()

    #drop useless columns and rows
    file = file.drop(columns = ['Description (Item)', 'Schedule (Comments)', 'Hub inventory', 'Vendor'])
    file = file[(file['Procurement type'] == 'B/S') | (file['Procurement type'] == 'Buysell')].reset_index(drop = True)

    #adjust qty columns name and units
    qtycol = file.filter(like='Single Shortage QTY (K)').columns.tolist()
    

    for i in qtycol:
        file[i] = file[i].apply(lambda x: x.upper().strip() if type(x) == str else x)
        file[i] = file[i].replace("NEW", 0)
        file[i] = file[i].replace("NEW ADD", 0)
        file[i] = file[i].apply(lambda x: x*1000)
    file = file.rename(columns= {qtycol[0]: 'Single Shortage QTY', qtycol[1]: 'Prev_Single Shortage QTY'})

    #replace ODM name
    file['ODM'] = file['ODM'].replace(ODMdict)
    return file   


# In[ ]:


target = Path (home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Single shortage')
PNFVPath = Path(home, 'HP Inc','GPSTW SOP - 2021 日新', 'PN FV description mapping table_ALL.xlsx')
ExternalReportFolder = Path(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'today')

PNFVFile = pd.read_excel(PNFVPath)
PNFVFile = PNFVFile [['PN', 'Descr']]
PNFVFile = PNFVFile.rename(columns = {'PN': 'HP PN'})


# In[ ]:


fileList = [str(x) for x in Path(target,'Archive').glob("*xlsx")]
fileListDateNum = [dt.strptime(x[-13:-5], "%Y%m%d").date() for x in fileList]

errorList = []
resultList = []


# In[ ]:


lookupSGdateList = []
i = 0
for _ in dateRange:
    if i == len(fileListDateNum)-1:
        lookupSGdateList.append(fileListDateNum[i])
        continue
    
    if _ < fileListDateNum[i+1]:
        lookupSGdateList.append(fileListDateNum[i])
    else:
        i = i + 1
        lookupSGdateList.append(fileListDateNum[i])


# In[ ]:


dateRange = [i.strftime("%Y%m%d") for i in dateRange]
lookupSGdateList = [i.strftime("%Y%m%d") for i in lookupSGdateList]


# In[ ]:


zip(dateRange, lookupSGdateList)


# dateList = result['reportDate'].tolist()
# max(dateList)

# In[ ]:


def addKey(res: pd.DataFrame) -> tuple[list, pd.DataFrame]:
    LatestSGMaterial = res
    LatestSGMaterial = LatestSGMaterial.merge(PNFVFile, on = 'HP PN', how = 'left')
    LatestSGMaterial['Key'] = LatestSGMaterial['ODM'] + LatestSGMaterial['Descr']
    KeyList = LatestSGMaterial['Key'].tolist()
    print("addKey done!")
    return KeyList, res


# ### concat current day external report

# In[ ]:


def concatExternal(extReportPathList: list, KeyList: list) -> pd.DataFrame:
    externalResultDFList = []
    if not extReportPathList:
        return 

    for _ in extReportPathList:
        try: 
            temp = pd.read_excel(_)
            temp['ODM'] = temp['ODM'].ffill()
            temp['ODM'] = temp['ODM'].replace(ODMdict)
            temp['FV/Des'] = temp['FV/Des'].ffill()
            #temp['ETA'] = temp['ETA'].ffill()
            temp['key'] = temp['ODM'] + temp['FV/Des']
            temp = temp[temp.key.isin(KeyList)]
            
            try:
                temp = temp[['ODM', 'FV/Des', 'HP_PN', 'ETA', 'GPS Remark']]
            except:
                temp = temp[['ODM', 'FV/Des', 'HP PN', 'ETA', 'GPS Remark']]
                temp = temp.rename(columns = {'HP PN' : 'HP_PN'})
                print("Rocky wrong format!")
                
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
    try:
        externalResultDF = pd.concat(externalResultDFList)
    except Exception as e:
        print(e)
        print("No single shortage match!")
        return pd.DataFrame(columns = ['ODM', 'FV/Des', 'HP_PN', 'ETA', 'GPS Remark'])
    print('External process done!')
    return externalResultDF


# ### lookup PNFV and merge external reportm
# ### output

# In[ ]:


def mergeNoutput(SGres: pd.DataFrame, extRes: pd.DataFrame, dateStr: str) -> None:
    SGres = SGres.merge(PNFVFile.rename(columns = {'PN': 'HP PN'}), on = 'HP PN', how = 'left')
    print(SGres[SGres['HP PN'] == 'M83662-N13'])
    SGres = SGres.merge(extRes.rename(columns = {'FV/Des' : 'Descr'}), on = ['ODM', 'Descr'], how = 'left')
    SGres = SGres.drop_duplicates()
    SGres.to_excel(Path(target, 'test','total singal shortage_' + dateStr +'.xls'), index = False, engine = 'xlwt')
    print("Output done!")
    return SGres


# In[ ]:


dateRange.reverse()


# In[ ]:


lookupSGdateList.reverse()


# In[ ]:


for i, j in zip(dateRange[0], lookupSGdateList[0]):
    print(i, j)


# In[ ]:


[f for f in glob.glob(str(Path(target, 'Single shortage ' + '*')))][-1]


# In[ ]:


et = today
# et = '20230512'
# sg = '20230209'

fname = [f for f in glob.glob(str(Path(target, 'Single shortage ' + '*')))][-1]
ExternalReport = [f for f in glob.glob(str(Path(ExternalReportFolder, et + '*')))]
if not ExternalReport:
    print("No external on " + et)
    #exit()

try:
    file = pd.read_excel(str(fname))
    sg_res = clean(str(fname), file, et)
    print(str(fname) + " process done!")
except Exception as e:
    errorList.append([str(fname), e])
    print(e)
    print(str(fname) + " process failed!")
    #exit()

k, sg_res = addKey(sg_res)
ext_res = concatExternal(ExternalReport, k)
if ext_res is None:
    print("No external on " + et)
    #exit()

sg_res = mergeNoutput(sg_res, ext_res, et)


# In[ ]:


conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()


# In[ ]:


#sg_res = sg_res.rename(columns={'HP PN' : 'HP_PN'})


# In[ ]:


sg_res['ETA'] = sg_res['ETA'].apply(lambda x: x.replace("'","") if type(x) == str else x)


# In[ ]:


sg_res


# In[ ]:


for index, row in sg_res.iterrows():
    sg_Commodity = row['Commodity']
    sg_Qty = row['Single Shortage QTY']
    sg_ODM = row['ODM']
    sg_series = row['Series']
    sg_PN_all = row['HP PN']
    sg_pQty = row['Prev_Single Shortage QTY']
    sg_pTyoe = row['Procurement type']
    sg_reportDate = row['reportDate']
    sg_ETA = row['ETA']
    sg_gpsRemark = row['GPS Remark']
    sg_lastSGreportDate = row['LastSGreportDate']
    #sg_PN_single = row['HP PN']


    cursor.execute(f"INSERT INTO CSI.OPS.GPS_tbl_ops_Single_shortage ( Commodity, [Single Shortage QTY], ODM, Series, [HP PN], [Prev_Single Shortage QTY], \
                    [Procurement type], reportDate, ETA, [GPS Remark], LastSGreportDate)\
                    VALUES('{sg_Commodity}','{sg_Qty}','{sg_ODM}','{sg_series}','{sg_PN_all}', '{sg_pQty}','{sg_pTyoe}','{sg_reportDate}',\
                           '{sg_ETA}','{sg_gpsRemark}', '{sg_lastSGreportDate}')".replace("'NaT'", "NULL").replace("'nan'", "NULL"))

conn.commit()


# In[ ]:


conn.close()


# In[ ]:


ExternalreportArchiveFolder = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'Archive')

for f in os.listdir(ExternalReportFolder):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(ExternalReportFolder, f), os.path.join(ExternalreportArchiveFolder, f))

for f in os.listdir(os.path.join(ExternalReportFolder, 'amend')):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(ExternalReportFolder, 'amend', f), os.path.join(ExternalreportArchiveFolder, 'amend', f))


# In[ ]:




