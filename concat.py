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


# ### Merge and clean FD

# #### Merge

# In[ ]:


# concat FD
FD_files = []
for FD_file in glob.glob(str(os.path.join(FD_folder,"*.xlsx"))):
    print(FD_file)
    FD = pd.read_excel(FD_file)
    FD_files.append(FD)
FD_all = pd.concat(FD_files)


# In[ ]:


# remember inplace = True
FD_all['len_FV'] = FD_all['FV'].str.len()
FD_all['len_platform'] = FD_all['Platform'].str.len()
FD_all = FD_all.sort_values( ['len_FV','len_platform'] , ascending = [False,False] )
FD_all.reset_index( drop = True, inplace = True )


# #### Move row with max len to the top row

# In[ ]:


# Index where B is longest
FD_idmax = FD_all['len_platform'].max()
df_max = FD_all.loc[FD_all['len_platform'] == FD_idmax]
df_max_to_add_FD = df_max.head(1)


# In[ ]:


# isin用在list
# drop the max len row
FD_all = FD_all.drop([FD_all.index[df_max_to_add_FD.index.values[0]]]).reset_index( drop = True )


# In[ ]:


# concat and put the max len row on the top
FD_concat = [df_max_to_add_FD,FD_all]
FD_output = pd.concat(FD_concat)
FD_output.reset_index( drop = True , inplace = True)


# In[ ]:


# cut more than 500
FD_output['Platform'] = FD_output['Platform'].apply(lambda x: x[:450] if len(x) > 500 else x)
FD_output = FD_output.drop( columns = ['len_FV','len_platform'])
FD_output['Item'] = FD_output['Item'].astype(str)


# ### Merge and clean Shortage

# #### Merge

# In[ ]:


# concat shortage
shortage_files = []
for shortage_file in glob.glob(str(os.path.join(shortage_folder,"*.xlsx"))):
    print(shortage_file)
    shortage = pd.read_excel(shortage_file)
    shortage_files.append(shortage)
shortage_all = pd.concat(shortage_files)


# In[ ]:


# create value to sort 
shortage_all['len_FV'] = shortage_all['FV'].str.len()
shortage_all['len_platform'] = shortage_all['Platform'].str.len()
shortage_all = shortage_all.sort_values(['len_FV','len_platform'] , ascending=[False,False])
shortage_all.reset_index( drop = True, inplace = True )


# In[ ]:


try:
    shortage_all['HP_PN'] = shortage_all['HP_PN'].apply(lambda x: x[:128] if len(x) > 128 else x)
except:    
    pass


# #### Move row with max len to the top row

# In[ ]:


# find the max length 
shortage_idmax = shortage_all['len_platform'].max()
shortage_max = shortage_all.loc[shortage_all['len_platform'] == shortage_idmax]
df_max_to_add_shortage = shortage_max.head(1)
df_max_to_add_shortage


# In[ ]:


# drop the max len row
shortage_all = shortage_all.drop([shortage_all.index[df_max_to_add_shortage.index.values[0]]]).reset_index( drop = True )


# In[ ]:


# concat and put the max len row on the top
shortage_concat = [ df_max_to_add_shortage , shortage_all ]
shortage_output = pd.concat(shortage_concat)
shortage_output.reset_index( drop = True , inplace = True)


# In[ ]:


# cut more than 500
shortage_output['Platform'] = shortage_output['Platform'].apply(lambda x: x[:450] if len(x) > 500 else x)
shortage_output = shortage_output.drop( columns= ['len_FV','len_platform'])
shortage_output['Item'] = shortage_output['Item'].astype(str)


# ### Merge and clean PNbasedDetail

# #### Merge

# In[ ]:


# concat PNbasedDetail
PNbasedDetail_files = []
for PNbasedDetail_file in glob.glob(str(os.path.join(PNbasedDetail_folder,"*.xlsx"))):
    print(PNbasedDetail_file)
    PNbasedDetail = pd.read_excel(PNbasedDetail_file)
    PNbasedDetail_files.append(PNbasedDetail)
PNbasedDetail_all = pd.concat(PNbasedDetail_files)


# In[ ]:


# create list to sort 
character_limit_list = ['GPS Remark','ODM use column1','ODM use column2','ODM use column3','ODM use column4','ODM use column5']
sort_list = []
asc_list = []
for _ in character_limit_list:

    try:
        PNbasedDetail_all[str('len_' + _)] = PNbasedDetail_all[_].str.len()
        sort_list.append(str('len_' + _))
        asc_list.append(True)
    except:    
        pass


# #### Move row with max len to the top row

# In[ ]:


# sort value
PNbasedDetail_all = PNbasedDetail_all.sort_values( sort_list , ascending = asc_list )
PNbasedDetail_all.reset_index( drop=True , inplace=True ) 


# In[ ]:


# find the max length & concat
PNbasedDetail_max_files = []

# a for loop to calculate all the len max
for i in range( 1 , len(sort_list) ):
    PNbasedDetail_idmax = PNbasedDetail_all [ sort_list[i] ].max()
    PNbasedDetail_max = PNbasedDetail_all.loc[ PNbasedDetail_all[ sort_list[i] ] == PNbasedDetail_idmax ]
    PNbasedDetail_max_files.append( PNbasedDetail_max.head(1) )
    df_max_to_add_PNbasedDetail_temp = pd.concat( PNbasedDetail_max_files )

# sometimes with duplicates 
df_max_to_add_PNbasedDetail = df_max_to_add_PNbasedDetail_temp.drop_duplicates()


# In[ ]:


# check point
df_max_to_add_PNbasedDetail


# In[ ]:


# drop the max len row
for i in range( 0, len(df_max_to_add_PNbasedDetail.index.values)):
    PNbasedDetail_all = PNbasedDetail_all.drop([PNbasedDetail_all.index[df_max_to_add_PNbasedDetail.index.values[i]]])


# In[ ]:


# concat and put on the top
PNbasedDetail_concat_list = [ df_max_to_add_PNbasedDetail , PNbasedDetail_all ]
PNbasedDetail_output = pd.concat(PNbasedDetail_concat_list).reset_index( drop = True )
PNbasedDetail_output


# In[ ]:


# cut more than 500
for _ in character_limit_list:
    try:
        PNbasedDetail_output[_] = PNbasedDetail_output[_].apply(lambda x: x[:450] if len(x) > 500 else x)
    except:    
        pass


# In[ ]:


# final step, drop calculate step and output
PNbasedDetail_output = PNbasedDetail_output.drop( columns = sort_list )
PNbasedDetail_output['Item'] = PNbasedDetail_output['Item'].astype(str)
PNbasedDetail_output


# ### Output concated FD, Shortage, and PNbasedDetail files

# In[ ]:


# apache airflow to upload SQL ( currently to desktop )
FD_output.to_excel(os.path.join(home, 'Desktop', 'FD_all.xlsx'), index=False)
shortage_output.to_excel(os.path.join(home, 'Desktop', 'Shortage_all.xlsx'), index=False)
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

