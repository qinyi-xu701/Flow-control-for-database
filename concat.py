#!/usr/bin/env python
# coding: utf-8

# ### Import Package and set path

# In[ ]:


import pandas as pd
import glob
import os
from datetime import date
import shutil


# In[ ]:


# home and time
home = os.path.expanduser("~")
todaystr = date.today().strftime('%Y-%m-%d')

# set up concat directories
targetFolder = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )')
FD_folder = os.path.join(targetFolder, "FD_today")
shortage_folder = os.path.join(targetFolder ,"shortage_today")
PNbasedDetail_folder = os.path.join(targetFolder ,"PNbasedDetail_today")


# ### Function Merge and Sort

# In[ ]:


def merge(path: str) -> pd.DataFrame:
    # concat
    temp_file_list = []
    for f in glob.glob(path):
        # print(f)
        temp_file = pd.read_excel(f)
        temp_file_list.append(temp_file)
    All = pd.concat(temp_file_list)
    
    return All


# In[ ]:


def maxLen(df_all: pd.DataFrame, sort_index: list) -> pd.DataFrame:
    # sort based on len
    sort_list = []
    for _ in sort_index:
        try:
            df_all[str('len_' + _)] = df_all[_].str.len()
            sort_list.append(str('len_' + _))
        except Exception as e:
            print(e)
    df_all = df_all.reset_index(drop = True)

    max_files = []
    for i, ele in enumerate(sort_list):
        idmax = df_all[ele].max()
        max = df_all[df_all[ele] == idmax]
        max_files.append(max.head(1))
    df_max_to_add = pd.concat(max_files).drop_duplicates()

    # drop the max len row
    for i, ele in enumerate(df_max_to_add.index.values):
        df_all = df_all.drop([df_all.index[ele]])

    # concat and put on the top
    output = pd.concat([df_max_to_add, df_all]).reset_index( drop = True )

    # cut more than 500
    for _ in sort_index:
        try:
            output[_] = output[_].apply(lambda x: x[:450] if len(x) > 500 else x)
        except Exception as e:
            print(e)
    
    # final step, drop calculate step and output
    output = output.drop(columns = sort_list)
    output['Item'] = output['Item'].astype(str)

    return output


# ### Generate FD, shortage, PNDetail table

# In[ ]:


FD = merge(str(os.path.join(FD_folder,"*.xlsx")))
FD_output = maxLen(FD, ['FV','Platform'])


shortage = merge(str(os.path.join(shortage_folder,"*.xlsx")))
try:
    shortage['HP_PN'] = shortage['HP_PN'].apply(lambda x: x[:128] if len(x) > 128 else x)
except:    
    pass
Shortage_output = maxLen(shortage, ['FV','Platform'])


PN = merge(str(os.path.join(PNbasedDetail_folder,"*.xlsx")))
PNbasedDetail_output = maxLen(PN, ['GPS Remark','ODM use column1','ODM use column2','ODM use column3','ODM use column4','ODM use column5'])


# ### Output concated FD, Shortage, and PNbasedDetail files

# In[ ]:


# apache airflow to upload SQL ( currently to desktop )
FD_output.to_excel(os.path.join(home, 'Desktop', 'FD_all.xlsx'), index=False)
Shortage_output.to_excel(os.path.join(home, 'Desktop', 'Shortage_all.xlsx'), index=False)
PNbasedDetail_output.to_excel(os.path.join(home, 'Desktop', 'PNbasedDetail_all.xlsx'), index=False)


# ### Move file to archive

# In[ ]:


FD_folder = os.path.join(targetFolder, "FD_today")
FD_archive_folder = os.path.join(targetFolder, 'FD_Archive_After_1025')

for f in os.listdir(FD_folder):
    shutil.move(os.path.join(FD_folder, f), os.path.join(FD_archive_folder, f))
    
shortage_folder = os.path.join(targetFolder ,"shortage_today")
shortage_archive_folder = os.path.join(targetFolder ,"Shortage_Archive_After_1025")

for f in os.listdir(shortage_folder):
    shutil.move(os.path.join(shortage_folder, f), os.path.join(shortage_archive_folder, f))

PNbasedDetail_folder = os.path.join(targetFolder ,"PNbasedDetail_today")
PNbasedDetail_archive_folder = os.path.join(targetFolder ,"PNbasedDetail_Archive_After_1025")

for f in os.listdir(PNbasedDetail_folder):
    shutil.move(os.path.join(PNbasedDetail_folder, f), os.path.join(PNbasedDetail_archive_folder, f))

