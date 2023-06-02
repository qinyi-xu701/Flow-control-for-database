#!/usr/bin/env python
# coding: utf-8

# #### Deal with re-upload or patching data

# In[ ]:


import re
import os
import glob
import pyodbc
import shutil
import datetime
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import date


# In[ ]:


# home and time
home = Path.home()
todaystr = date.today().strftime('%Y-%m-%d')
targetFolder = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )')

# file path
FD_path =  os.path.join(Path(home, targetFolder, 'FD_today','amend'))
Shortage_path =  os.path.join(Path(home, targetFolder, 'Shortage_today','amend'))
PNbasedDetail_path =  os.path.join(Path(home, targetFolder, 'PNbasedDetail_today','amend'))


# In[ ]:


# connect SQL server
conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()


# #### Important maxlen function

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

    df_max_to_add.index.values.sort()

    # drop the max len row
    for i, ele in enumerate(df_max_to_add.index.values):
        df_all = df_all.drop([df_all.index[ele - i]])

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


# #### Delete old FD and upload new FD

# In[ ]:


# Get the names of all files in the folder with .xlsx extension
file_names = [file_name for file_name in os.listdir(FD_path) if file_name.endswith('.xlsx')]

for file in file_names:
    condition = file.split('_')
    print(condition)

    # date condition
    date_condition = datetime.datetime(int(condition[0][:4]),int(condition[0][4:6]),int(condition[0][6:]))
    # commodity condition
    commodity_condition = (condition[1])
    # ODM condition
    ODM_condition = (condition[2])
    # buyer condition
    buyer_condition = (condition[3])

    # delete statement
    delete_query = "DELETE FROM CSI.OPS.GPS_tbl_ops_fd WHERE ReportDate = ? AND Commodity = ? AND ODM = ? AND BuyerName = ?"
    params = (date_condition, commodity_condition, ODM_condition, buyer_condition)
    cursor.execute(delete_query, params)

    # check result
    print(f"{cursor.rowcount} rows deleted from CSI.OPS.GPS_tbl_ops_fd")


# In[ ]:


# Create an empty list to store the fd
fd_amend_temp = []

# Loop through all the Excel files in the folder
for fd in glob.glob(os.path.join(FD_path, '*.xlsx')):
    # Read the data from the Excel file into a pandas dataframe
    dff = pd.read_excel(fd)
    # Append the dataframe to the list
    fd_amend_temp.append(dff)

# Concatenate all the dataframes into a single dataframe
FD_amend_data_temp = pd.concat(fd_amend_temp, ignore_index=True)

# Max len check
# FD_amend_data = maxLen(FD_amend_data_temp, ['FV','Platform'])
FD_amend_data = FD_amend_data_temp.copy()

for index, row in FD_amend_data.iterrows():
    f_ODM = row['ODM']
    f_Item = row['Item']
    f_Commodity = row['Commodity']
    f_FV = row['FV']
    f_Platform = row['Platform']
    f_HP_PN = row['HP_PN']
    f_Supplier = row['Supplier']
    f_PN = row['HP PN']
    f_ReportDate = row['ReportDate']
    f_FDdate = row['FDdate']
    f_FDQty = row['FDQty']
    f_BuyerName = row['BuyerName']
    
    cursor.execute(f"INSERT INTO CSI.OPS.GPS_tbl_ops_fd ( ODM,Item,Commodity,FV,Platform,Supplier,[HP PN],FDdate,FDQty,Reportdate,BuyerName )\
                    VALUES('{f_ODM}','{f_Item}','{f_Commodity}','{f_FV}','{f_Platform}','{f_Supplier}','{f_PN}','{f_FDdate}','{f_FDQty}','{f_ReportDate}','{f_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()


# #### Delete old shortage and upload new shortage

# In[ ]:


# Get the names of all files in the folder with .xlsx extension
file_names = [file_name for file_name in os.listdir(Shortage_path) if file_name.endswith('.xlsx')]

for file in file_names:
    condition = file.split('_')
    print(condition)

    # date condition
    date_condition = datetime.datetime(int(condition[0][:4]),int(condition[0][4:6]),int(condition[0][6:]))
    # commodity condition
    commodity_condition = (condition[1])
    # ODM condition
    ODM_condition = (condition[2])
    # buyer condition
    buyer_condition = (condition[3])

    # delete statement
    delete_query = "DELETE FROM CSI.OPS.GPS_tbl_ops_shortage_ext WHERE ReportDate = ? AND Commodity = ? AND ODM = ? AND BuyerName = ?"
    params = (date_condition, commodity_condition, ODM_condition, buyer_condition)
    cursor.execute(delete_query, params)

    # check result
    print(f"{cursor.rowcount} rows deleted from CSI.OPS.GPS_tbl_ops_shortage_ext" )


# In[ ]:


# Create an empty list to store the Shortage
Shortage_amend_temp = []

# Loop through all the Excel files in the folder
for Shortage in glob.glob(os.path.join(Shortage_path, '*.xlsx')):
    # Read the data from the Excel file into a pandas dataframe
    dfs = pd.read_excel(Shortage)
    # Append the dataframe to the list
    Shortage_amend_temp.append(dfs)

# Concatenate all the dataframes into a single dataframe
Shortage_amend_data_temp = pd.concat(Shortage_amend_temp, ignore_index=True)

# Max len check
try:
    Shortage_amend_data_temp['HP_PN'] = Shortage_amend_data_temp['HP_PN'].apply(lambda x: x[:128] if len(x) > 128 else x)
except:    
    pass

Shortage_amend_data = maxLen(Shortage_amend_data_temp, ['Platform','FV'])


for index, row in Shortage_amend_data.iterrows():
    s_ODM = row['ODM']
    s_Item = row['Item']
    s_Commodity = row['Commodity']
    s_FV = row['FV']
    s_Platform = row['Platform']
    s_P1 = row['P1']
    s_P2 = row['Net P2']
    s_P3 = row['Net P3']
    s_Total = row['Total Shortage Qty']
    s_BT = row['BT shortage']
    s_working = row['Working on upside']
    s_ReportDate = pd.to_datetime(row['ReportDate'])
    s_lastFDdate = pd.to_datetime(row['last FD date'])
    s_BuyerName = row['BuyerName']
    s_PN = row['HP_PN']

    cursor.execute(f"INSERT INTO CSI.OPS.GPS_tbl_ops_shortage_ext ( ODM,Item,Commodity,FV,Platform,P1,[Net P2],[Net P3],[Total Shortage Qty],[BT shortage],[Working on upside],ReportDate,[last FD date],HP_PN,BuyerName )\
                    VALUES('{s_ODM}','{s_Item}','{s_Commodity}','{s_FV}','{s_Platform}','{s_P1}','{s_P2}','{s_P3}','{s_Total}','{s_BT}','{s_working}','{s_ReportDate}','{s_lastFDdate}','{s_PN}','{s_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()


# #### Delete old PNbasedDetail and upload new PNbasedDetail

# In[ ]:


# Get the names of all files in the folder with .xlsx extension
file_names = [file_name for file_name in os.listdir(PNbasedDetail_path) if file_name.endswith('.xlsx')]

for file in file_names:
    condition = file.split('_')
    print(condition)

    # date condition
    date_condition = datetime.datetime(int(condition[0][:4]),int(condition[0][4:6]),int(condition[0][6:]))
    # commodity condition
    commodity_condition = (condition[1])
    # ODM condition
    ODM_condition = (condition[2])
    # buyer condition
    buyer_condition = (condition[3])

    # delete statement
    delete_query = "DELETE FROM CSI.OPS.GPS_tbl_ops_PNbasedDetail WHERE ReportDate = ? AND Commodity = ? AND ODM = ? AND BuyerName = ?"
    params = (date_condition, commodity_condition, ODM_condition, buyer_condition)
    cursor.execute(delete_query, params)

    # check result
    print(f"{cursor.rowcount} rows deleted from CSI.OPS.GPS_tbl_ops_PNbasedDetail" )

conn.commit()


# In[ ]:


# Create an empty list to store the PNbasedDetail
PNbasedDetail_amend_temp = []

# Loop through all the Excel files in the folder
for PNbasedDetail in glob.glob(os.path.join(PNbasedDetail_path, '*.xlsx')):
    # Read the data from the Excel file into a pandas dataframe
    dfp = pd.read_excel(PNbasedDetail)
    # Append the dataframe to the list
    PNbasedDetail_amend_temp.append(dfp)

# Concatenate all the dataframes into a single dataframe
PNbasedDetail_amend_data_temp = pd.concat(PNbasedDetail_amend_temp, ignore_index=True)

# fillna for PNbasedDetail
for i in ['GPS Remark', 'ODM use column1','ODM use column2','ODM use column3','ODM use column4','ODM use column5']:
    PNbasedDetail_amend_data_temp[i] = PNbasedDetail_amend_data_temp[i].fillna("")

# Max len check
# PNbasedDetail_amend_data = maxLen(PNbasedDetail_amend_data_temp, ['GPS Remark','ODM use column1','ODM use column2','ODM use column3','ODM use column4','ODM use column5'])



for index, row in PNbasedDetail_amend_data_temp.iterrows():
    p_ODM = row['ODM']
    p_Item = row['Item']
    p_Commodity = row['Commodity']
    p_PN = row['HP PN']
    p_Remark = row['GPS Remark'].replace("\'", "\'\'")
    p_stock = row['852 stock']
    p_change = row['852 stock change']
    p_over = row['Over pull qty']
    p_ODM1 = str(row['ODM use column1']).replace("\'", "\'\'")
    p_ODM2 = str(row['ODM use column2']).replace("\'", "\'\'")
    p_ODM3 = str(row['ODM use column3']).replace("\'", "\'\'")
    p_ODM4 = str(row['ODM use column4']).replace("\'", "\'\'")
    p_ODM5 = str(row['ODM use column5']).replace("\'", "\'\'")
    p_ReportDate = row['ReportDate']
    p_BuyerName = row['BuyerName']

    cursor.execute(f"INSERT INTO CSI.OPS.GPS_tbl_ops_PNbasedDetail ( ODM,Item,Commodity,[HP PN],[GPS Remark],[852 stock],[852 stock change],[Over pull qty],\
                    [ODM use column1],[ODM use column2],[ODM use column3],[ODM use column4],[ODM use column5],ReportDate,BuyerName )\
                    VALUES('{p_ODM}','{p_Item}','{p_Commodity}','{p_PN}','{p_Remark}','{p_stock}','{p_change}','{p_over}','{p_ODM1}','{p_ODM2}','{p_ODM3}','{p_ODM4}','{p_ODM5}','{p_ReportDate}','{p_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()
conn.close()


# #### Move uploaded data to archieve

# In[ ]:


# FD Archieve folder define
FD_archive_folder = os.path.join(targetFolder, 'FD_Archive_After_1025')
# Move FD
for f in os.listdir(FD_path):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(FD_path, f), os.path.join(FD_archive_folder, f))
    else:
        pass
    
# Shortage Archieve folder define
shortage_archive_folder = os.path.join(targetFolder ,"Shortage_Archive_After_1025")
# Move Shortage
for f in os.listdir(Shortage_path):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(Shortage_path, f), os.path.join(shortage_archive_folder, f))
    else:
        pass

# PNbasedDetail Archieve folder define
PNbasedDetail_archive_folder = os.path.join(targetFolder ,"PNbasedDetail_Archive_After_1025")
# Move PNbasedDetail
for f in os.listdir(PNbasedDetail_path):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(PNbasedDetail_path, f), os.path.join(PNbasedDetail_archive_folder, f))
    else:
        pass

PNbasedDetail_archive_folder = os.path.join(targetFolder ,"PNbasedDetail_Archive_After_1025")
# Move PNbasedDetail
for f in os.listdir(PNbasedDetail_path):
    if f.endswith('.xlsx'):
        shutil.move(os.path.join(PNbasedDetail_path, f), os.path.join(PNbasedDetail_archive_folder, f))
    else:
        pass

